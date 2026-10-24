"""Small, dependency-light Excel formula evaluator (openpyxl only).

Why it exists: the checker must re-run a model with inputs changed (to prove a
sensitivity axis feeds its output, and to recompute Excel Data Tables). The
`formulas` library cannot load workbooks that contain Data Tables, and
LibreOffice is not always installed, so this engine does the substitution work.

Usage:
    from xleval import Model
    m = Model("model.xlsx")
    ctx = m.context()                      # base case
    ctx.value("Outputs", "D12")
    ctx2 = m.context({("Inputs", 5, 5): 0.2})   # (sheet, row, col) overrides
"""
from __future__ import annotations

import math
import re
import sys
import threading
from dataclasses import dataclass, field

import openpyxl
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

sys.setrecursionlimit(200000)


class XLErr:
    __slots__ = ("code",)

    def __init__(self, code):
        self.code = code

    def __repr__(self):
        return self.code

    def __eq__(self, other):
        return isinstance(other, XLErr) and other.code == self.code

    def __hash__(self):
        return hash(self.code)


DIV0, VALUE, REF, NAME, NA, NUM, CIRC = (XLErr(c) for c in
                                         ("#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NUM!", "#CIRC!"))

CELL_RE = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$")
COL_RE = re.compile(r"^\$?([A-Z]{1,3})$")
ROW_RE = re.compile(r"^\$?(\d+)$")


def split_sheet(text):
    """'Sheet 1'!A1:B2 -> ('Sheet 1', 'A1:B2'); A1 -> (None, 'A1')."""
    if "!" not in text:
        return None, text
    sheet, addr = text.rsplit("!", 1)
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, addr


def parse_cell(a):
    m = CELL_RE.match(a.upper())
    if not m:
        return None
    return int(m.group(2)), column_index_from_string(m.group(1))


def addr(r, c):
    return f"{get_column_letter(c)}{r}"


@dataclass
class DataTable:
    sheet: str
    r1: int  # result block bounds
    c1: int
    r2: int
    c2: int
    two_d: bool
    row_oriented: bool  # dtr
    in1: tuple | None  # (row, col) of input cell r1
    in2: tuple | None

    def contains(self, r, c):
        return self.r1 <= r <= self.r2 and self.c1 <= c <= self.c2

    @property
    def ref(self):
        return f"{addr(self.r1, self.c1)}:{addr(self.r2, self.c2)}"

    def corner(self):
        return (self.r1 - 1, self.c1 - 1)

    def plan(self, r, c):
        """Return (formula_cell, [(input_cell, axis_cell), ...]) for result cell (r, c)."""
        if self.two_d:
            # r1 takes the values of the top row, r2 the values of the left column
            return self.corner(), [(self.in1, (self.r1 - 1, c)), (self.in2, (r, self.c1 - 1))]
        if self.row_oriented:
            return (r, self.c1 - 1), [(self.in1, (self.r1 - 1, c))]
        return (self.r1 - 1, c), [(self.in1, (r, self.c1 - 1))]

    def axis_cells(self):
        if self.two_d:
            top = [(self.r1 - 1, c) for c in range(self.c1, self.c2 + 1)]
            left = [(r, self.c1 - 1) for r in range(self.r1, self.r2 + 1)]
            return top, left
        if self.row_oriented:
            return [(self.r1 - 1, c) for c in range(self.c1, self.c2 + 1)], []
        return [], [(r, self.c1 - 1) for r in range(self.r1, self.r2 + 1)]


class Rng:
    """A rectangular reference, evaluated lazily."""
    __slots__ = ("ctx", "sheet", "r1", "c1", "r2", "c2")

    def __init__(self, ctx, sheet, r1, c1, r2, c2):
        self.ctx, self.sheet = ctx, sheet
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2

    @property
    def shape(self):
        return self.r2 - self.r1 + 1, self.c2 - self.c1 + 1

    def grid(self):
        return [[self.ctx.get((self.sheet, r, c)) for c in range(self.c1, self.c2 + 1)]
                for r in range(self.r1, self.r2 + 1)]

    def flat(self):
        return [v for row in self.grid() for v in row]

    def scalar(self):
        if self.shape == (1, 1):
            return self.ctx.get((self.sheet, self.r1, self.c1))
        return VALUE


# ----------------------------------------------------------------------------- parser
class Parser:
    BIN = {"=": 10, "<>": 10, "<": 10, ">": 10, "<=": 10, ">=": 10, "&": 20,
           "+": 30, "-": 30, "*": 40, "/": 40, "^": 50}
    PREFIX_BP = 60

    def __init__(self, formula):
        tok = Tokenizer(formula)
        self.toks = [t for t in tok.items if t.type != Token.WSPACE]
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else None

    def next(self):
        t = self.peek()
        self.i += 1
        return t

    def parse(self):
        node = self.expr(0)
        return node

    def expr(self, min_bp):
        t = self.next()
        if t is None:
            return ("missing",)
        if t.type == Token.OP_PRE:
            operand = self.expr(self.PREFIX_BP)
            left = ("un", t.value, operand)
        elif t.type == Token.PAREN and t.subtype == Token.OPEN:
            left = self.expr(0)
            self.next()  # closing paren
        elif t.type == Token.FUNC and t.subtype == Token.OPEN:
            name = t.value[:-1].upper().replace("_XLFN.", "").replace("_XLWS.", "")
            args = []
            nt = self.peek()
            if nt is not None and nt.type == Token.FUNC and nt.subtype == Token.CLOSE:
                self.next()
            else:
                while True:
                    nt = self.peek()
                    if nt is not None and nt.type == Token.SEP:
                        args.append(("missing",))
                    elif nt is not None and nt.type == Token.FUNC and nt.subtype == Token.CLOSE:
                        args.append(("missing",))
                    else:
                        args.append(self.expr(0))
                    nt = self.next()
                    if nt is None or (nt.type == Token.FUNC and nt.subtype == Token.CLOSE):
                        break
            left = ("fn", name, args)
        elif t.type == Token.OPERAND:
            left = self.operand(t)
        else:
            left = ("missing",)
        while True:
            t = self.peek()
            if t is None:
                break
            if t.type == Token.OP_POST:
                self.next()
                left = ("pct", left)
                continue
            if t.type == Token.OP_IN and t.value in self.BIN:
                bp = self.BIN[t.value]
                if bp <= min_bp and not (t.value == "^" and bp == min_bp and False):
                    break
                self.next()
                right = self.expr(bp)
                left = ("bin", t.value, left, right)
                continue
            if t.type == Token.OP_IN and t.value == ":":
                # range built from two references, e.g. A1:INDEX(...): rare, approximate with left
                self.next()
                right = self.expr(55)
                left = ("span", left, right)
                continue
            break
        return left

    @staticmethod
    def operand(t):
        v = t.value
        if t.subtype == Token.NUMBER:
            return ("num", float(v))
        if t.subtype == Token.TEXT:
            return ("str", v[1:-1].replace('""', '"'))
        if t.subtype == Token.LOGICAL:
            return ("bool", v.upper() == "TRUE")
        if t.subtype == Token.ERROR:
            return ("err", XLErr(v))
        return ("refname", v)


# ----------------------------------------------------------------------------- model
class Model:
    def __init__(self, path, cached_path=None):
        self.path = path
        self.wb = openpyxl.load_workbook(path, data_only=False)
        self.wbv = openpyxl.load_workbook(cached_path or path, data_only=True)
        self.sheetnames = self.wb.sheetnames
        self._lower = {s.lower(): s for s in self.sheetnames}
        self.cells = {}
        self.cached = {}
        self.datatables: list[DataTable] = []
        self.max_rc = {}
        for ws in self.wb.worksheets:
            wv = self.wbv[ws.title] if ws.title in self.wbv.sheetnames else None
            mr = mc = 0
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if v is None:
                        continue
                    key = (ws.title, c.row, c.column)
                    if isinstance(v, ArrayFormula):
                        v = v.text if v.text.startswith("=") else "=" + v.text
                    self.cells[key] = v
                    mr, mc = max(mr, c.row), max(mc, c.column)
                    if isinstance(v, DataTableFormula):
                        self._add_dt(ws.title, v)
            if wv is not None:
                for row in wv.iter_rows():
                    for c in row:
                        if c.value is not None:
                            self.cached[(ws.title, c.row, c.column)] = c.value
                            mr, mc = max(mr, c.row), max(mc, c.column)
            self.max_rc[ws.title] = (mr, mc)
        self.names = {}
        try:
            for name, dn in self.wb.defined_names.items():
                self.names[name.upper()] = dn.attr_text
        except Exception:
            pass
        for ws in self.wb.worksheets:  # sheet-scoped names
            try:
                for name, dn in ws.defined_names.items():
                    self.names.setdefault(name.upper(), dn.attr_text)
            except Exception:
                pass
        self.dt_index = {}
        for dt in self.datatables:
            for r in range(dt.r1, dt.r2 + 1):
                for c in range(dt.c1, dt.c2 + 1):
                    self.dt_index[(dt.sheet, r, c)] = dt
        self._ast = {}
        self._prec = None

    def _add_dt(self, sheet, f):
        a, _, b = f.ref.partition(":")
        r1, c1 = parse_cell(a)
        r2, c2 = parse_cell(b or a)
        truthy = lambda x: str(x).lower() in ("1", "true")
        in1 = parse_cell(f.r1) if f.r1 else None
        in2 = parse_cell(f.r2) if getattr(f, "r2", None) else None
        self.datatables.append(DataTable(sheet, r1, c1, r2, c2, truthy(f.dt2D), truthy(f.dtr), in1, in2))

    # -- helpers
    def sheet(self, name):
        if name is None:
            return None
        return self._lower.get(name.lower())

    def is_formula(self, key):
        v = self.cells.get(key)
        return isinstance(v, str) and v.startswith("=")

    def formula(self, key):
        v = self.cells.get(key)
        return v if isinstance(v, str) and v.startswith("=") else None

    def ast(self, key):
        if key not in self._ast:
            f = self.cells[key]
            try:
                self._ast[key] = Parser(f).parse()
            except Exception:
                self._ast[key] = ("err", NAME)
        return self._ast[key]

    def resolve(self, text, cur_sheet):
        """Reference text -> (sheet, r1, c1, r2, c2) or None."""
        if text.upper() in self.names:
            text = self.names[text.upper()]
            if "," in text:
                text = text.split(",")[0]
        sheet, a = split_sheet(text)
        if sheet and sheet.startswith("["):
            return None  # external workbook
        sh = self.sheet(sheet) if sheet else cur_sheet
        if sh is None:
            return None
        parts = a.replace("$", "").upper().split(":")
        if len(parts) == 1:
            rc = parse_cell(parts[0])
            if not rc:
                return None
            return (sh, rc[0], rc[1], rc[0], rc[1])
        p, q = parts[0], parts[1]
        rc1, rc2 = parse_cell(p), parse_cell(q)
        if rc1 and rc2:
            return (sh, min(rc1[0], rc2[0]), min(rc1[1], rc2[1]), max(rc1[0], rc2[0]), max(rc1[1], rc2[1]))
        mr, mc = self.max_rc.get(sh, (1, 1))
        if COL_RE.match(p) and COL_RE.match(q):
            return (sh, 1, column_index_from_string(p), max(mr, 1), column_index_from_string(q))
        if ROW_RE.match(p) and ROW_RE.match(q):
            return (sh, int(p), 1, int(q), max(mc, 1))
        return None

    def refs_in(self, key):
        """Static references of a formula cell: list of (sheet, r1, c1, r2, c2)."""
        out = []

        def walk(n):
            if n[0] == "refname":
                r = self.resolve(n[1], key[0])
                if r:
                    out.append(r)
            elif n[0] in ("fn",):
                for a in n[2]:
                    walk(a)
            elif n[0] == "bin":
                walk(n[2]); walk(n[3])
            elif n[0] == "span":
                walk(n[1]); walk(n[2])
            elif n[0] in ("un",):
                walk(n[2])
            elif n[0] == "pct":
                walk(n[1])
        walk(self.ast(key))
        return out

    def precedents(self):
        """key -> set of direct precedent keys (ranges expanded, data tables included)."""
        if self._prec is not None:
            return self._prec
        prec = {}
        for key in self.cells:
            if self.is_formula(key):
                s = set()
                for (sh, r1, c1, r2, c2) in self.refs_in(key):
                    if (r2 - r1 + 1) * (c2 - c1 + 1) > 20000:
                        continue
                    for r in range(r1, r2 + 1):
                        for c in range(c1, c2 + 1):
                            s.add((sh, r, c))
                prec[key] = s
        for dt in self.datatables:
            for r in range(dt.r1, dt.r2 + 1):
                for c in range(dt.c1, dt.c2 + 1):
                    f, pairs = dt.plan(r, c)
                    s = {(dt.sheet,) + f} | {(dt.sheet,) + ax for _, ax in pairs}
                    s |= {(dt.sheet,) + inp for inp, _ in pairs if inp}
                    prec[(dt.sheet, r, c)] = s
        self._prec = prec
        return prec

    def dependents(self):
        dep = {}
        for k, s in self.precedents().items():
            for p in s:
                dep.setdefault(p, set()).add(k)
        return dep

    def closure(self, key, stop_at_overrides=False):
        """All transitive precedents of key."""
        prec = self.precedents()
        seen, stack = set(), [key]
        while stack:
            k = stack.pop()
            for p in prec.get(k, ()):
                if p not in seen:
                    seen.add(p)
                    stack.append(p)
        return seen

    def context(self, overrides=None):
        return Ctx(self, overrides or {})


def run_big_stack(fn, *a, **kw):
    """Run fn in a thread with a large stack (deep formula chains)."""
    res = {}

    def target():
        try:
            res["v"] = fn(*a, **kw)
        except BaseException as e:  # noqa
            res["e"] = e
    threading.stack_size(512 * 1024 * 1024)
    t = threading.Thread(target=target)
    t.start()
    t.join()
    if "e" in res:
        raise res["e"]
    return res.get("v")


# ----------------------------------------------------------------------------- evaluation
def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def to_num(v):
    if isinstance(v, XLErr):
        return v
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if is_num(v):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(",", ""))
        except ValueError:
            return VALUE
    if hasattr(v, "year"):  # dates are numbers in Excel; good enough for checks
        return VALUE
    return VALUE


def to_bool(v):
    if isinstance(v, XLErr):
        return v
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    if is_num(v):
        return v != 0
    if isinstance(v, str):
        if v.upper() in ("TRUE", "FALSE"):
            return v.upper() == "TRUE"
        return VALUE
    return VALUE


class Ctx:
    def __init__(self, model: Model, overrides):
        self.m = model
        self.ov = {}
        for k, v in overrides.items():
            self.ov[(model.sheet(k[0]) or k[0], k[1], k[2])] = v
        self.memo = {}
        self.busy = set()
        self.unsupported = set()

    # public
    def value(self, sheet, a):
        r, c = parse_cell(a)
        return self.get((self.m.sheet(sheet), r, c))

    def get(self, key):
        if key in self.ov:
            return self.ov[key]
        if key in self.memo:
            return self.memo[key]
        raw = self.m.cells.get(key)
        dt = self.m.dt_index.get(key)
        if dt is not None:
            self._compute_dt(dt)
            return self.memo.get(key)
        if isinstance(raw, str) and raw.startswith("="):
            if key in self.busy:
                return CIRC
            self.busy.add(key)
            try:
                v = self.ev(self.m.ast(key), key[0])
                if isinstance(v, Rng):
                    v = v.scalar() if v.shape == (1, 1) else v.grid()[0][0]
                if isinstance(v, list):
                    v = v[0][0] if v and isinstance(v[0], list) else (v[0] if v else None)
                if v is None:
                    v = 0.0
            finally:
                self.busy.discard(key)
            self.memo[key] = v
            return v
        if isinstance(raw, DataTableFormula):
            return None
        v = raw
        if isinstance(v, int) and not isinstance(v, bool):
            v = float(v)
        self.memo[key] = v
        return v

    def _compute_dt(self, dt: DataTable):
        sh = dt.sheet
        for r in range(dt.r1, dt.r2 + 1):
            for c in range(dt.c1, dt.c2 + 1):
                fcell, pairs = dt.plan(r, c)
                ov = dict(self.ov)
                for inp, ax in pairs:
                    if inp is None:
                        continue
                    ov[(sh,) + inp] = self.get((sh,) + ax)
                sub = Ctx(self.m, {})
                sub.ov = ov
                self.memo[(sh, r, c)] = sub.get((sh,) + fcell)
                self.unsupported |= sub.unsupported

    # core
    def ev(self, n, sheet):
        t = n[0]
        if t == "num":
            return n[1]
        if t == "str":
            return n[1]
        if t == "bool":
            return n[1]
        if t == "err":
            return n[1]
        if t == "missing":
            return None
        if t == "refname":
            r = self.m.resolve(n[1], sheet)
            if r is None:
                if n[1].upper() in ("TRUE", "FALSE"):
                    return n[1].upper() == "TRUE"
                return REF if "!" in n[1] or CELL_RE.match(n[1].replace("$", "")) else NAME
            return Rng(self, *r)
        if t == "span":
            a = self.ev(n[1], sheet)
            b = self.ev(n[2], sheet)
            if isinstance(a, Rng) and isinstance(b, Rng):
                return Rng(self, a.sheet, min(a.r1, b.r1), min(a.c1, b.c1), max(a.r2, b.r2), max(a.c2, b.c2))
            return REF
        if t == "un":
            v = self.scalarize(self.ev(n[2], sheet))
            if isinstance(v, list):
                return self.elementwise(lambda x: self._neg(x, n[1]), v)
            return self._neg(v, n[1])
        if t == "pct":
            v = to_num(self.scalarize(self.ev(n[1], sheet)))
            return v if isinstance(v, XLErr) else v / 100.0
        if t == "bin":
            a = self.ev(n[2], sheet)
            b = self.ev(n[3], sheet)
            return self.binop(n[1], a, b)
        if t == "fn":
            return self.call(n[1], n[2], sheet)
        return VALUE

    @staticmethod
    def _neg(v, op):
        x = to_num(v)
        if isinstance(x, XLErr):
            return x
        return -x if op == "-" else x

    def scalarize(self, v):
        if isinstance(v, Rng):
            if v.shape == (1, 1):
                return v.scalar()
            return v.grid()
        return v

    @staticmethod
    def elementwise(f, grid):
        return [[f(x) for x in row] for row in grid]

    def binop(self, op, a, b):
        a, b = self.scalarize(a), self.scalarize(b)
        if isinstance(a, list) or isinstance(b, list):
            ga = a if isinstance(a, list) else None
            gb = b if isinstance(b, list) else None
            rows = len(ga or gb)
            cols = len((ga or gb)[0])
            out = []
            for i in range(rows):
                row = []
                for j in range(cols):
                    x = ga[min(i, len(ga) - 1)][min(j, len(ga[0]) - 1)] if ga else a
                    y = gb[min(i, len(gb) - 1)][min(j, len(gb[0]) - 1)] if gb else b
                    row.append(self.binop(op, x, y))
                out.append(row)
            return out
        if isinstance(a, XLErr):
            return a
        if isinstance(b, XLErr):
            return b
        if op == "&":
            return f"{self.text(a)}{self.text(b)}"
        if op in ("=", "<>", "<", ">", "<=", ">="):
            return self.compare(op, a, b)
        x, y = to_num(a), to_num(b)
        if isinstance(x, XLErr):
            return x
        if isinstance(y, XLErr):
            return y
        try:
            if op == "+":
                return x + y
            if op == "-":
                return x - y
            if op == "*":
                return x * y
            if op == "/":
                return DIV0 if y == 0 else x / y
            if op == "^":
                if x == 0 and y < 0:
                    return DIV0
                r = x ** y
                return NUM if isinstance(r, complex) else float(r)
        except OverflowError:
            return NUM
        return VALUE

    @staticmethod
    def text(v):
        if v is None:
            return ""
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if is_num(v):
            return ("%.15g" % v)
        return str(v)

    @staticmethod
    def compare(op, a, b):
        def norm(v):
            if v is None:
                return (0, 0.0)
            if isinstance(v, bool):
                return (2, v)
            if is_num(v):
                return (0, float(v))
            return (1, str(v).lower())
        if a is None and isinstance(b, str):
            a = ""
        if b is None and isinstance(a, str):
            b = ""
        ta, tb = norm(a), norm(b)
        if ta[0] != tb[0]:
            res = {"=": False, "<>": True}.get(op)
            if res is not None:
                return res
            return {"<": ta[0] < tb[0], ">": ta[0] > tb[0], "<=": ta[0] <= tb[0], ">=": ta[0] >= tb[0]}[op]
        x, y = ta[1], tb[1]
        if op == "=":
            return x == y if not isinstance(x, float) else abs(x - y) <= 1e-12 * max(1.0, abs(x), abs(y))
        if op == "<>":
            return not Ctx.compare("=", a, b)
        return {"<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y}[op]

    # functions --------------------------------------------------------------
    def nums(self, args, sheet, errors=True):
        """Flatten args to numbers, Excel aggregate style."""
        out = []
        for a in args:
            v = self.ev(a, sheet)
            if isinstance(v, Rng):
                for x in v.flat():
                    if isinstance(x, XLErr):
                        return x
                    if is_num(x):
                        out.append(float(x))
            elif isinstance(v, list):
                for row in v:
                    for x in row:
                        if isinstance(x, XLErr):
                            return x
                        if is_num(x):
                            out.append(float(x))
                        elif isinstance(x, bool):
                            out.append(1.0 if x else 0.0)
            else:
                if v is None and a[0] == "missing":
                    continue
                x = to_num(v)
                if isinstance(x, XLErr):
                    return x
                out.append(x)
        return out

    def arg(self, args, i, sheet, default=None):
        if i >= len(args) or args[i][0] == "missing":
            return default
        return self.scalarize(self.ev(args[i], sheet))

    def grid_of(self, node, sheet):
        v = self.ev(node, sheet)
        if isinstance(v, Rng):
            return v.grid()
        if isinstance(v, list):
            return v
        return [[v]]

    def call(self, name, args, sheet):
        f = FUNCS.get(name)
        if f is None:
            self.unsupported.add(name)
            return NAME
        try:
            return f(self, args, sheet)
        except (ZeroDivisionError,):
            return DIV0
        except (ValueError, OverflowError, IndexError, TypeError):
            return VALUE


def _agg(fn):
    def w(ctx, args, sheet):
        xs = ctx.nums(args, sheet)
        if isinstance(xs, XLErr):
            return xs
        return fn(xs)
    return w


def _num1(fn):
    def w(ctx, args, sheet):
        x = to_num(ctx.arg(args, 0, sheet))
        if isinstance(x, XLErr):
            return x
        try:
            r = fn(x)
        except (ValueError, ZeroDivisionError):
            return NUM
        return r
    return w


def f_if(ctx, args, sheet):
    c = to_bool(ctx.arg(args, 0, sheet))
    if isinstance(c, XLErr):
        return c
    if c:
        return ctx.arg(args, 1, sheet, True)
    return ctx.arg(args, 2, sheet, False)


def f_iferror(ctx, args, sheet):
    v = ctx.arg(args, 0, sheet)
    return ctx.arg(args, 1, sheet) if isinstance(v, XLErr) else v


def f_and(ctx, args, sheet):
    vals = []
    for a in args:
        v = ctx.scalarize(ctx.ev(a, sheet))
        items = [x for row in v for x in row] if isinstance(v, list) else [v]
        for x in items:
            b = to_bool(x)
            if isinstance(b, XLErr):
                return b
            vals.append(b)
    return vals


def f_round(mode):
    def w(ctx, args, sheet):
        x = to_num(ctx.arg(args, 0, sheet))
        d = to_num(ctx.arg(args, 1, sheet, 0))
        if isinstance(x, XLErr):
            return x
        d = int(d)
        m = 10.0 ** d
        if mode == "round":
            return math.floor(abs(x) * m + 0.5) / m * (1 if x >= 0 else -1)
        if mode == "up":
            return math.ceil(abs(x) * m - 1e-12) / m * (1 if x >= 0 else -1)
        return math.floor(abs(x) * m + 1e-12) / m * (1 if x >= 0 else -1)
    return w


def f_log(ctx, args, sheet):
    x = to_num(ctx.arg(args, 0, sheet))
    b = to_num(ctx.arg(args, 1, sheet, 10.0))
    if isinstance(x, XLErr):
        return x
    if x <= 0:
        return NUM
    return math.log(x) / math.log(b)


def f_npv(ctx, args, sheet):
    r = to_num(ctx.arg(args, 0, sheet))
    if isinstance(r, XLErr):
        return r
    xs = ctx.nums(args[1:], sheet)
    if isinstance(xs, XLErr):
        return xs
    return sum(v / (1 + r) ** (i + 1) for i, v in enumerate(xs))


def irr(xs, guess=0.1):
    if not xs or all(x >= 0 for x in xs) or all(x <= 0 for x in xs):
        return NUM

    def npv(r):
        return sum(v / (1 + r) ** i for i, v in enumerate(xs))
    r = guess
    for _ in range(100):
        try:
            f = npv(r)
            d = sum(-i * v / (1 + r) ** (i + 1) for i, v in enumerate(xs))
        except (OverflowError, ZeroDivisionError):
            break
        if d == 0:
            break
        nr = r - f / d
        if abs(nr - r) < 1e-12:
            return nr
        r = nr
        if r <= -0.9999:
            break
    lo, hi = -0.9999, 10.0
    flo, fhi = npv(lo), npv(hi)
    if flo * fhi > 0:
        return NUM
    for _ in range(300):
        mid = (lo + hi) / 2
        fm = npv(mid)
        if flo * fm <= 0:
            hi, fhi = mid, fm
        else:
            lo, flo = mid, fm
    return (lo + hi) / 2


def f_irr(ctx, args, sheet):
    xs = ctx.nums(args[:1], sheet)
    if isinstance(xs, XLErr):
        return xs
    g = to_num(ctx.arg(args, 1, sheet, 0.1))
    return irr(xs, g)


def f_sumproduct(ctx, args, sheet):
    grids = [ctx.grid_of(a, sheet) for a in args]
    shape = (len(grids[0]), len(grids[0][0]))
    if any((len(g), len(g[0])) != shape for g in grids):
        return VALUE
    total = 0.0
    for i in range(shape[0]):
        for j in range(shape[1]):
            p = 1.0
            for g in grids:
                x = g[i][j]
                if isinstance(x, XLErr):
                    return x
                if isinstance(x, bool):
                    x = 1.0 if x else 0.0
                p *= float(x) if is_num(x) else 0.0
            total += p
    return total


def f_index(ctx, args, sheet):
    g = ctx.grid_of(args[0], sheet)
    r = ctx.arg(args, 1, sheet, 0)
    c = ctx.arg(args, 2, sheet, None)
    r = int(to_num(r)) if r is not None else 0
    if c is None:
        if len(g) == 1:
            return g[0][r - 1] if r >= 1 else VALUE
        if len(g[0]) == 1:
            return g[r - 1][0] if r >= 1 else VALUE
        c = 1
    c = int(to_num(c))
    if r < 1 or c < 1 or r > len(g) or c > len(g[0]):
        return REF
    return g[r - 1][c - 1]


def f_match(ctx, args, sheet):
    x = ctx.arg(args, 0, sheet)
    g = ctx.grid_of(args[1], sheet)
    items = [row[0] for row in g] if len(g[0]) == 1 else g[0]
    mt = to_num(ctx.arg(args, 2, sheet, 1.0))
    if mt == 0:
        for i, v in enumerate(items):
            if Ctx.compare("=", v, x) is True and (is_num(v) == is_num(x)):
                return float(i + 1)
        return NA
    best = None
    for i, v in enumerate(items):
        if v is None or isinstance(v, XLErr):
            continue
        if (mt > 0 and Ctx.compare("<=", v, x)) or (mt < 0 and Ctx.compare(">=", v, x)):
            best = i + 1
    return float(best) if best else NA


def f_choose(ctx, args, sheet):
    i = to_num(ctx.arg(args, 0, sheet))
    if isinstance(i, XLErr):
        return i
    i = int(i)
    if i < 1 or i >= len(args):
        return VALUE
    return ctx.arg(args, i, sheet)


def _criteria(crit):
    if is_num(crit):
        return lambda v: is_num(v) and abs(float(v) - crit) <= 1e-12 * max(1, abs(crit))
    s = str(crit)
    m = re.match(r"^(<=|>=|<>|<|>|=)?(.*)$", s)
    op, rhs = m.group(1) or "=", m.group(2)
    try:
        rv = float(rhs)
    except ValueError:
        rv = rhs
    return lambda v: Ctx.compare(op, v, rv) is True


def f_countif(ctx, args, sheet):
    g = ctx.grid_of(args[0], sheet)
    pred = _criteria(ctx.arg(args, 1, sheet))
    return float(sum(1 for row in g for v in row if pred(v)))


def f_sumif(ctx, args, sheet):
    g = ctx.grid_of(args[0], sheet)
    pred = _criteria(ctx.arg(args, 1, sheet))
    s = ctx.grid_of(args[2], sheet) if len(args) > 2 else g
    tot = 0.0
    for i, row in enumerate(g):
        for j, v in enumerate(row):
            if pred(v):
                x = s[i][j]
                if is_num(x):
                    tot += x
    return tot


def f_rank(ctx, args, sheet):
    x = to_num(ctx.arg(args, 0, sheet))
    g = ctx.grid_of(args[1], sheet)
    order = to_num(ctx.arg(args, 2, sheet, 0.0))
    xs = [float(v) for row in g for v in row if is_num(v)]
    if isinstance(x, XLErr):
        return x
    if order == 0:
        return float(1 + sum(1 for v in xs if v > x))
    return float(1 + sum(1 for v in xs if v < x))


def f_rows(ctx, args, sheet):
    v = ctx.ev(args[0], sheet)
    return float(v.shape[0]) if isinstance(v, Rng) else 1.0


def f_columns(ctx, args, sheet):
    v = ctx.ev(args[0], sheet)
    return float(v.shape[1]) if isinstance(v, Rng) else 1.0


def f_isnumber(ctx, args, sheet):
    v = ctx.arg(args, 0, sheet)
    return is_num(v)


def f_isblank(ctx, args, sheet):
    v = ctx.arg(args, 0, sheet)
    return v is None


def f_iserror(ctx, args, sheet):
    return isinstance(ctx.arg(args, 0, sheet), XLErr)


def f_count(ctx, args, sheet):
    n = 0
    for a in args:
        v = ctx.ev(a, sheet)
        items = v.flat() if isinstance(v, Rng) else [v]
        n += sum(1 for x in items if is_num(x))
    return float(n)


def f_counta(ctx, args, sheet):
    n = 0
    for a in args:
        v = ctx.ev(a, sheet)
        items = v.flat() if isinstance(v, Rng) else [v]
        n += sum(1 for x in items if x is not None and x != "")
    return float(n)


def f_power(ctx, args, sheet):
    return ctx.binop("^", ctx.arg(args, 0, sheet), ctx.arg(args, 1, sheet))


def f_mod(ctx, args, sheet):
    x, y = to_num(ctx.arg(args, 0, sheet)), to_num(ctx.arg(args, 1, sheet))
    if y == 0:
        return DIV0
    return x - y * math.floor(x / y)


def f_concat(ctx, args, sheet):
    return "".join(Ctx.text(ctx.arg(args, i, sheet)) for i in range(len(args)))


def _median(xs):
    s = sorted(xs)
    n = len(s)
    if n == 0:
        return NUM
    return s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2


def pct_inc(xs, k):
    """Excel PERCENTILE / PERCENTILE.INC (linear interpolation)."""
    s = sorted(xs)
    if not s or not 0 <= k <= 1:
        return NUM
    h = (len(s) - 1) * k
    lo = math.floor(h)
    return s[lo] + (h - lo) * (s[min(lo + 1, len(s) - 1)] - s[lo])


def f_percentile(ctx, args, sheet):
    xs = ctx.nums(args[:1], sheet)
    k = to_num(ctx.arg(args, 1, sheet))
    for v in (xs, k):
        if isinstance(v, XLErr):
            return v
    return pct_inc(xs, k)


def _bool_agg(kind):
    def w(ctx, args, sheet):
        v = f_and(ctx, args, sheet)
        if isinstance(v, XLErr):
            return v
        return all(v) if kind == "and" else any(v)
    return w


FUNCS = {
    "SUM": _agg(sum),
    "AVERAGE": _agg(lambda xs: sum(xs) / len(xs) if xs else DIV0),
    "MIN": _agg(lambda xs: min(xs) if xs else 0.0),
    "MAX": _agg(lambda xs: max(xs) if xs else 0.0),
    "PRODUCT": _agg(lambda xs: math.prod(xs) if xs else 0.0),
    "MEDIAN": _agg(_median),
    "PERCENTILE": f_percentile,
    "PERCENTILE.INC": f_percentile,
    "COUNT": f_count,
    "COUNTA": f_counta,
    "IF": f_if,
    "IFERROR": f_iferror,
    "AND": _bool_agg("and"),
    "OR": _bool_agg("or"),
    "NOT": lambda ctx, a, s: (lambda b: b if isinstance(b, XLErr) else not b)(to_bool(ctx.arg(a, 0, s))),
    "TRUE": lambda ctx, a, s: True,
    "FALSE": lambda ctx, a, s: False,
    "ABS": _num1(abs),
    "INT": _num1(lambda x: float(math.floor(x))),
    "TRUNC": _num1(lambda x: float(math.trunc(x))),
    "SIGN": _num1(lambda x: float((x > 0) - (x < 0))),
    "LN": _num1(math.log),
    "LOG10": _num1(math.log10),
    "LOG": f_log,
    "EXP": _num1(math.exp),
    "SQRT": _num1(math.sqrt),
    "POWER": f_power,
    "MOD": f_mod,
    "PI": lambda ctx, a, s: math.pi,
    "ROUND": f_round("round"),
    "ROUNDUP": f_round("up"),
    "ROUNDDOWN": f_round("down"),
    "NPV": f_npv,
    "IRR": f_irr,
    "SUMPRODUCT": f_sumproduct,
    "INDEX": f_index,
    "MATCH": f_match,
    "CHOOSE": f_choose,
    "COUNTIF": f_countif,
    "SUMIF": f_sumif,
    "RANK": f_rank,
    "RANK.EQ": f_rank,
    "ROWS": f_rows,
    "COLUMNS": f_columns,
    "ISNUMBER": f_isnumber,
    "ISBLANK": f_isblank,
    "ISERROR": f_iserror,
    "CONCATENATE": f_concat,
    "CONCAT": f_concat,
}
