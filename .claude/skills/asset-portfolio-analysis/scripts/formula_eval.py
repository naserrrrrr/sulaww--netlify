"""Mini Excel formula evaluator used by inspect_workbook.py for consistency checks.

Scope (deliberately small, deterministic, stdlib only):
  * numbers, strings, booleans, error literals, parentheses, unary +/-, %, ^ * / + - & and comparisons
  * A1 references with optional sheet prefix ('Sheet name'!A1, Sheet!$A$1:$B$9, A:A, 3:3)
  * structured references: Table[col], Table[[col]], Table[[#This Row],[col]], Table[@col],
    Table[@[col]], Table[[#Totals],[col]], Table[[#Data],[col]], Table[[#All],[col]],
    Table[[#Headers],[col]], Table[[colA]:[colB]], Table[] and bare [@col] inside a table
  * functions: SUM, SUBTOTAL (1/101 average, 2/102 count, 3/103 counta, 4/104 max, 5/105 min,
    9/109 sum), COUNTA, COUNT, AVERAGE, MIN, MAX, PRODUCT, ROUND/ROUNDUP/ROUNDDOWN, ABS, INT, IF,
    IFERROR, NA

Anything else raises NotEvaluable (the caller counts it as "not evaluated").
The evaluator never writes to the workbook. Precedent values come from the cached values
stored in the file; a precedent formula without a cached value is evaluated recursively.

The "book" object passed to Evaluator must provide:
  book.sheet(name)            -> sheet view or None (case/space tolerant lookup is the caller's job)
  book.table(name)            -> TableMeta-like or None
  book.table_at(sheet, r, c)  -> TableMeta-like containing the cell, or None
and a sheet view must provide:
  .name, .max_row, .max_col, .has_filter
  .value(r, c)      -> literal value or cached value of a formula cell (None if empty)
  .formula(r, c)    -> formula text ("=...") or None
  .is_hidden_row(r) -> bool
A TableMeta-like must provide: .name, .sheet, .min_row, .max_row, .min_col, .max_col,
  .header_rows (0/1), .totals_rows (0/1), .columns (list of column names in order).
"""
from __future__ import annotations

import datetime as _dt
import math
import re
from typing import Any, Dict, List, Optional, Tuple

ERRORS = ("#GETTING_DATA", "#DIV/0!", "#VALUE!", "#SPILL!", "#CALC!", "#NAME?", "#NULL!", "#NUM!",
          "#REF!", "#N/A")
TOLERANCE = 0.5
MAX_COL = 16384
MAX_ROW = 1048576


class NotEvaluable(Exception):
    """Raised when a formula uses something outside the supported subset."""


class XLError:
    __slots__ = ("code",)

    def __init__(self, code: str):
        self.code = code

    def __eq__(self, other: Any) -> bool:
        return isinstance(other, XLError) and other.code == self.code

    def __hash__(self) -> int:
        return hash(self.code)

    def __repr__(self) -> str:
        return self.code


# ---------------------------------------------------------------------------
# Column helpers
# ---------------------------------------------------------------------------
def col_to_idx(s: str) -> int:
    n = 0
    for ch in s.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def idx_to_col(n: int) -> str:
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


# ---------------------------------------------------------------------------
# Tokenizer
# ---------------------------------------------------------------------------
_REF_RE = re.compile(
    r"\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?"   # A1 or A1:B2
    r"|\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})"                      # A:B
    r"|\$?(\d+):\$?(\d+)"                                          # 3:9
)
_WORD_RE = re.compile(r"[\w.$\\]+")
_NUM_RE = re.compile(r"(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?")
_WORD_CHARS = re.compile(r"[\w.$\\]")


class Area:
    """A rectangular reference. r1/r2 or c1/c2 may be None for whole columns / rows."""
    __slots__ = ("sheet", "r1", "c1", "r2", "c2", "struct_cols")

    def __init__(self, sheet, r1, c1, r2, c2, struct_cols=None):
        self.sheet, self.r1, self.c1, self.r2, self.c2 = sheet, r1, c1, r2, c2
        self.struct_cols = struct_cols  # list of table column names when built from a structured ref

    def is_single(self) -> bool:
        return self.r1 is not None and self.r1 == self.r2 and self.c1 is not None and self.c1 == self.c2

    def __repr__(self) -> str:
        return f"Area({self.sheet!r},{self.r1},{self.c1},{self.r2},{self.c2})"


def _boundary_ok(s: str, j: int) -> bool:
    return j >= len(s) or not (_WORD_CHARS.match(s[j]) or s[j] in "([!")


def _match_ref(s: str, i: int, sheet: Optional[str]):
    m = _REF_RE.match(s, i)
    if not m or not _boundary_ok(s, m.end()):
        return None, i
    g = m.groups()
    if g[0] is not None:
        c1, r1 = col_to_idx(g[0]), int(g[1])
        if g[2] is not None:
            c2, r2 = col_to_idx(g[2]), int(g[3])
        else:
            c2, r2 = c1, r1
        if c1 > MAX_COL or c2 > MAX_COL or r1 < 1 or r2 < 1:
            return None, i
        return ("REF", {"sheet": sheet, "r1": min(r1, r2), "c1": min(c1, c2), "r2": max(r1, r2), "c2": max(c1, c2)}), m.end()
    if g[4] is not None:
        c1, c2 = col_to_idx(g[4]), col_to_idx(g[5])
        return ("REF", {"sheet": sheet, "r1": None, "c1": min(c1, c2), "r2": None, "c2": max(c1, c2)}), m.end()
    r1, r2 = int(g[6]), int(g[7])
    return ("REF", {"sheet": sheet, "r1": min(r1, r2), "c1": None, "r2": max(r1, r2), "c2": None}), m.end()


def _read_brackets(s: str, i: int) -> Tuple[str, int]:
    """s[i] == '['. Return (inner text, index after the matching ']'). Honors ' escapes."""
    depth = 0
    j = i
    n = len(s)
    while j < n:
        ch = s[j]
        if ch == "'" and j + 1 < n:
            j += 2
            continue
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
        j += 1
    raise NotEvaluable("unbalanced brackets")


def tokenize(formula: str) -> List[tuple]:
    s = formula[1:] if formula.startswith("=") else formula
    i, n = 0, len(s)
    out: List[tuple] = []
    while i < n:
        ch = s[i]
        if ch in " \t\r\n":
            i += 1
            continue
        if ch == '"':
            j, buf = i + 1, []
            while True:
                if j >= n:
                    raise NotEvaluable("unterminated string")
                if s[j] == '"':
                    if j + 1 < n and s[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(s[j])
                j += 1
            out.append(("STR", "".join(buf)))
            i = j + 1
            continue
        if ch == "'":
            j, buf = i + 1, []
            while True:
                if j >= n:
                    raise NotEvaluable("unterminated sheet name")
                if s[j] == "'":
                    if j + 1 < n and s[j + 1] == "'":
                        buf.append("'")
                        j += 2
                        continue
                    break
                buf.append(s[j])
                j += 1
            j += 1
            if j >= n or s[j] != "!":
                raise NotEvaluable("quoted name without '!'")
            sheet = "".join(buf)
            if "]" in sheet:
                raise NotEvaluable("external workbook reference")
            if s.startswith("#REF!", j + 1):
                out.append(("ERR", "#REF!"))
                i = j + 6
                continue
            tok, k = _match_ref(s, j + 1, sheet)
            if tok is None:
                raise NotEvaluable("bad sheet reference")
            out.append(tok)
            i = k
            continue
        if ch == "#":
            up = s[i:i + 14].upper()
            for e in ERRORS:
                if up.startswith(e):
                    out.append(("ERR", e))
                    i += len(e)
                    break
            else:
                raise NotEvaluable("unknown # token")
            continue
        if ch.isdigit() or (ch == "." and i + 1 < n and s[i + 1].isdigit()):
            if ch.isdigit():
                tok, k = _match_ref(s, i, None)  # row range like 3:9
                if tok is not None and tok[1]["c1"] is None:
                    out.append(tok)
                    i = k
                    continue
            m = _NUM_RE.match(s, i)
            out.append(("NUM", float(m.group(0))))
            i = m.end()
            continue
        if ch == "[":
            content, j = _read_brackets(s, i)
            out.append(("STRUCT", {"table": None, "spec": content}))
            i = j
            continue
        if ch in "+-*/^&%":
            out.append(("OP", ch))
            i += 1
            continue
        if ch in "<>=":
            two = s[i:i + 2]
            if two in ("<>", "<=", ">="):
                out.append(("OP", two))
                i += 2
            else:
                out.append(("OP", ch))
                i += 1
            continue
        if ch == "(":
            out.append(("LP", ch))
            i += 1
            continue
        if ch == ")":
            out.append(("RP", ch))
            i += 1
            continue
        if ch in ",;":
            out.append(("SEP", ch))
            i += 1
            continue
        if ch in "{}":
            raise NotEvaluable("array constant")
        m = _WORD_RE.match(s, i)
        if not m:
            raise NotEvaluable(f"unexpected character {ch!r}")
        word, j = m.group(0), m.end()
        if j < n and s[j] == "!":
            if "[" in word or "]" in word:
                raise NotEvaluable("external workbook reference")
            if s.startswith("#REF!", j + 1):
                out.append(("ERR", "#REF!"))
                i = j + 6
                continue
            tok, k = _match_ref(s, j + 1, word)
            if tok is None:
                raise NotEvaluable("bad sheet reference")
            out.append(tok)
            i = k
            continue
        if j < n and s[j] == "(":
            name = word.upper()
            for pre in ("_XLFN._XLWS.", "_XLFN.", "_XLWS."):
                if name.startswith(pre):
                    name = name[len(pre):]
            out.append(("FUNC", name))
            i = j
            continue
        if j < n and s[j] == "[":
            content, k = _read_brackets(s, j)
            out.append(("STRUCT", {"table": word, "spec": content}))
            i = k
            continue
        up = word.upper()
        if up in ("TRUE", "FALSE"):
            out.append(("BOOL", up == "TRUE"))
            i = j
            continue
        tok, k = _match_ref(s, i, None)
        if tok is not None:
            out.append(tok)
            i = k
            continue
        out.append(("NAME", word))
        i = j
    return out


# ---------------------------------------------------------------------------
# Structured reference specification
# ---------------------------------------------------------------------------
def _unescape(name: str) -> str:
    out, i = [], 0
    while i < len(name):
        if name[i] == "'" and i + 1 < len(name):
            out.append(name[i + 1])
            i += 2
            continue
        out.append(name[i])
        i += 1
    return "".join(out)


def _split_items(spec: str) -> Tuple[List[str], List[str]]:
    items, seps, i, n = [], [], 0, len(spec)
    while i < n:
        ch = spec[i]
        if ch in " ":
            i += 1
            continue
        if ch == "[":
            inner, j = _read_brackets(spec, i)
            items.append(inner)
            i = j
            continue
        if ch in ",:":
            seps.append(ch)
            i += 1
            continue
        raise NotEvaluable("bad structured reference")
    return items, seps


_SPECIALS = {"#this row": "this_row", "#data": "data", "#totals": "totals", "#headers": "headers", "#all": "all"}


def parse_struct_spec(spec: str) -> Tuple[str, Optional[List[str]], bool]:
    """Return (area, columns, is_span). area in this_row|data|totals|headers|all."""
    if spec.strip() == "":
        return "data", None, False
    if spec.startswith("@"):
        rest = spec[1:]
        if rest.startswith("["):
            items, seps = _split_items(rest)
            return "this_row", [_unescape(x) for x in items], ":" in seps
        return "this_row", [_unescape(rest)], False
    if not spec.lstrip().startswith("["):
        low = spec.strip().lower()
        if low in _SPECIALS:
            return _SPECIALS[low], None, False
        return "data", [_unescape(spec)], False
    items, seps = _split_items(spec)
    area = None
    cols: List[str] = []
    for it in items:
        low = it.strip().lower()
        if low in _SPECIALS:
            a = _SPECIALS[low]
            if area is None:
                area = a
            elif {area, a} == {"headers", "data"} or {area, a} == {"data", "totals"}:
                area = "all"
            else:
                area = a
        else:
            cols.append(_unescape(it))
    is_span = ":" in seps and len(cols) == 2
    return area or "data", (cols or None), is_span


# ---------------------------------------------------------------------------
# Parser (precedence climbing) -> AST tuples
# ---------------------------------------------------------------------------
class _Parser:
    def __init__(self, toks: List[tuple]):
        self.t = toks
        self.i = 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def parse(self):
        node = self.comparison()
        if self.i != len(self.t):
            raise NotEvaluable("trailing tokens")
        return node

    def comparison(self):
        node = self.concat()
        while self.peek()[0] == "OP" and self.peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            node = ("bin", op, node, self.concat())
        return node

    def concat(self):
        node = self.additive()
        while self.peek() == ("OP", "&"):
            self.take()
            node = ("bin", "&", node, self.additive())
        return node

    def additive(self):
        node = self.term()
        while self.peek()[0] == "OP" and self.peek()[1] in ("+", "-"):
            op = self.take()[1]
            node = ("bin", op, node, self.term())
        return node

    def term(self):
        node = self.power()
        while self.peek()[0] == "OP" and self.peek()[1] in ("*", "/"):
            op = self.take()[1]
            node = ("bin", op, node, self.power())
        return node

    def power(self):
        node = self.unary()
        while self.peek() == ("OP", "^"):
            self.take()
            node = ("bin", "^", node, self.unary())
        return node

    def unary(self):
        tok = self.peek()
        if tok[0] == "OP" and tok[1] in ("-", "+"):
            self.take()
            inner = self.unary()
            return ("neg", inner) if tok[1] == "-" else ("pos", inner)
        return self.postfix()

    def postfix(self):
        node = self.primary()
        while self.peek() == ("OP", "%"):
            self.take()
            node = ("pct", node)
        return node

    def primary(self):
        kind, val = self.take()
        if kind == "NUM":
            return ("num", val)
        if kind == "STR":
            return ("str", val)
        if kind == "BOOL":
            return ("bool", val)
        if kind == "ERR":
            return ("err", val)
        if kind == "REF":
            return ("ref", val)
        if kind == "STRUCT":
            return ("struct", val)
        if kind == "NAME":
            return ("name", val)
        if kind == "LP":
            node = self.comparison()
            if self.take()[0] != "RP":
                raise NotEvaluable("missing ')'")
            return node
        if kind == "FUNC":
            if self.take()[0] != "LP":
                raise NotEvaluable("function without '('")
            args = []
            if self.peek()[0] == "RP":
                self.take()
                return ("func", val, args)
            while True:
                if self.peek()[0] in ("SEP", "RP"):
                    args.append(("missing",))
                else:
                    args.append(self.comparison())
                k = self.take()[0]
                if k == "RP":
                    break
                if k != "SEP":
                    raise NotEvaluable("bad argument list")
            return ("func", val, args)
        raise NotEvaluable(f"unexpected token {kind}")


def parse(formula: str):
    return _Parser(tokenize(formula)).parse()


# ---------------------------------------------------------------------------
# AST inspection helpers (used by the anomaly detectors)
# ---------------------------------------------------------------------------
def iter_ref_nodes(ast) -> List[tuple]:
    """All ('ref', dict) and ('struct', dict) nodes in the AST, in order."""
    out: List[tuple] = []

    def walk(node):
        k = node[0]
        if k in ("ref", "struct"):
            out.append(node)
        elif k == "bin":
            walk(node[2])
            walk(node[3])
        elif k in ("neg", "pos", "pct"):
            walk(node[1])
        elif k == "func":
            for a in node[2]:
                walk(a)

    walk(ast)
    return out


def functions_used(ast) -> List[str]:
    out: List[str] = []

    def walk(node):
        k = node[0]
        if k == "func":
            out.append(node[1])
            for a in node[2]:
                walk(a)
        elif k == "bin":
            walk(node[2])
            walk(node[3])
        elif k in ("neg", "pos", "pct"):
            walk(node[1])

    walk(ast)
    return out


_SUBTOTAL_KIND = {1: "average", 2: "count", 3: "counta", 4: "max", 5: "min", 9: "sum"}


def aggregate_kind(ast) -> Optional[Tuple[str, List[tuple]]]:
    """If the formula is a single aggregate call, return (kind, ref nodes); kind in sum|count|counta|average|min|max."""
    if ast[0] != "func":
        return None
    name, args = ast[1], ast[2]
    refs = [a for a in args if a[0] in ("ref", "struct")]
    if name == "SUBTOTAL" and args and args[0][0] == "num":
        fn = int(args[0][1]) % 100
        kind = _SUBTOTAL_KIND.get(fn)
        if kind is None:
            return None
        return kind, [a for a in args[1:] if a[0] in ("ref", "struct")]
    mapping = {"SUM": "sum", "COUNTA": "counta", "COUNT": "count", "AVERAGE": "average", "MIN": "min", "MAX": "max"}
    if name in mapping and refs and len(refs) == len(args):
        return mapping[name], refs
    return None


def additive_terms(ast) -> Optional[List[Tuple[int, tuple]]]:
    """If the AST is a pure +/- combination of references, return [(sign, ref_node), ...]."""
    out: List[Tuple[int, tuple]] = []

    def walk(node, sign) -> bool:
        k = node[0]
        if k in ("ref", "struct"):
            out.append((sign, node))
            return True
        if k == "bin" and node[1] in ("+", "-"):
            return walk(node[2], sign) and walk(node[3], sign if node[1] == "+" else -sign)
        if k == "pos":
            return walk(node[1], sign)
        if k == "neg":
            return walk(node[1], -sign)
        return False

    return out if walk(ast, 1) else None


# ---------------------------------------------------------------------------
# Evaluation
# ---------------------------------------------------------------------------
def _is_num(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _to_number(v: Any):
    """Excel-like coercion for arithmetic. Returns float or XLError."""
    if v is None:
        return 0.0
    if isinstance(v, XLError):
        return v
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        t = v.strip().replace(",", "")
        if t == "":
            return XLError("#VALUE!")
        try:
            return float(t)
        except ValueError:
            return XLError("#VALUE!")
    if isinstance(v, (_dt.date, _dt.datetime, _dt.time)):
        raise NotEvaluable("date arithmetic")
    raise NotEvaluable(f"unsupported value type {type(v).__name__}")


def _to_text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return str(v)


def _as_error(v: Any) -> Optional[XLError]:
    if isinstance(v, XLError):
        return v
    if isinstance(v, str) and v.strip().upper() in ERRORS:
        return XLError(v.strip().upper())
    return None


class _Ctx:
    __slots__ = ("sheet", "row", "col", "remap", "depth")

    def __init__(self, sheet, row, col, remap, depth):
        self.sheet, self.row, self.col, self.remap, self.depth = sheet, row, col, remap, depth


class Evaluator:
    def __init__(self, book, budget: int = 4_000_000, max_depth: int = 150):
        self.book = book
        self.budget = budget
        self.max_depth = max_depth
        self.cells_read = 0
        self.memo: Dict[tuple, Any] = {}
        self.visiting: set = set()
        self._parse_cache: Dict[str, Any] = {}

    # -- parsing -----------------------------------------------------------
    def parse(self, text: str):
        hit = self._parse_cache.get(text)
        if hit is not None:
            if isinstance(hit, NotEvaluable):
                raise hit
            return hit
        try:
            ast = parse(text)
        except NotEvaluable as exc:
            self._parse_cache[text] = exc
            raise
        except (IndexError, ValueError, TypeError) as exc:
            err = NotEvaluable(f"parse error: {exc}")
            self._parse_cache[text] = err
            raise err from exc
        if len(self._parse_cache) < 200_000:
            self._parse_cache[text] = ast
        return ast

    # -- references --------------------------------------------------------
    def resolve(self, node, sheet: str, row: int, col: int) -> Area:
        """Resolve a ('ref'|'struct', spec) node to an Area (no remap applied)."""
        kind, spec = node
        if kind == "ref":
            sh_name = spec["sheet"] or sheet
            sh = self.book.sheet(sh_name)
            if sh is None:
                raise NotEvaluable(f"unknown sheet {sh_name!r}")
            r1, r2, c1, c2 = spec["r1"], spec["r2"], spec["c1"], spec["c2"]
            if r1 is None:
                r1, r2 = 1, max(1, sh.max_row)
            if c1 is None:
                c1, c2 = 1, max(1, sh.max_col)
            return Area(sh.name, r1, c1, r2, c2)
        # structured
        tname = spec["table"]
        if tname is None:
            t = self.book.table_at(sheet, row, col)
        else:
            t = self.book.table(tname)
        if t is None:
            raise NotEvaluable(f"unknown table {tname!r}")
        area, cols, is_span = parse_struct_spec(spec["spec"])
        first_data = t.min_row + (t.header_rows or 0)
        last_data = t.max_row - (t.totals_rows or 0)
        if area == "data":
            r1, r2 = first_data, last_data
        elif area == "this_row":
            if not (t.min_row <= row <= t.max_row):
                raise NotEvaluable("#This Row outside table")
            r1 = r2 = row
        elif area == "totals":
            if not t.totals_rows:
                raise NotEvaluable("table has no totals row")
            r1 = r2 = t.max_row
        elif area == "headers":
            if not t.header_rows:
                raise NotEvaluable("table has no header row")
            r1 = r2 = t.min_row
        else:
            r1, r2 = t.min_row, t.max_row
        if cols is None:
            c1, c2 = t.min_col, t.max_col
            names = list(t.columns)
        else:
            idx = [self._table_col(t, c) for c in cols]
            if is_span:
                c1, c2 = min(idx), max(idx)
            elif len(idx) == 1:
                c1 = c2 = idx[0]
            else:
                raise NotEvaluable("multi-column structured reference")
            names = [t.columns[k - t.min_col] for k in range(c1, c2 + 1)]
        return Area(t.sheet, r1, c1, r2, c2, struct_cols=names)

    @staticmethod
    def _table_col(t, name: str) -> int:
        cols = list(t.columns)
        if name in cols:
            return t.min_col + cols.index(name)
        squeeze = lambda x: re.sub(r"\s+", " ", str(x or "")).strip().lower()
        target = squeeze(name)
        for k, c in enumerate(cols):
            if squeeze(c) == target:
                return t.min_col + k
        raise NotEvaluable(f"unknown table column {name!r}")

    # -- cell values -------------------------------------------------------
    def cell_value(self, sheet: str, r: int, c: int, depth: int = 0):
        self.cells_read += 1
        if self.cells_read > self.budget:
            raise NotEvaluable("evaluation budget exhausted")
        sh = self.book.sheet(sheet)
        if sh is None:
            raise NotEvaluable("unknown sheet")
        v = sh.value(r, c)
        f = sh.formula(r, c)
        if f is None or v is not None:
            err = _as_error(v)
            return err if err is not None else v
        key = (sh.name, r, c)
        if key in self.memo:
            return self.memo[key]
        if key in self.visiting:
            raise NotEvaluable("circular reference")
        if depth >= self.max_depth:
            raise NotEvaluable("precedent chain too deep")
        self.visiting.add(key)
        try:
            ast = self.parse(f)
            res = self._eval(ast, _Ctx(sh.name, r, c, None, depth + 1))
            res = self._scalar(res, _Ctx(sh.name, r, c, None, depth + 1))
        finally:
            self.visiting.discard(key)
        self.memo[key] = res
        return res

    def _area_cells(self, area: Area, ctx: _Ctx, skip_hidden=False, skip_subtotals=False):
        sh = self.book.sheet(area.sheet)
        if sh is None:
            raise NotEvaluable("unknown sheet")
        r2 = min(area.r2, max(sh.max_row, area.r1))
        c2 = min(area.c2, max(sh.max_col, area.c1))
        remap = ctx.remap or {}
        for r in range(area.r1, r2 + 1):
            if skip_hidden and sh.is_hidden_row(r):
                continue
            for c in range(area.c1, c2 + 1):
                cc = remap.get((sh.name, c), c)
                if skip_subtotals:
                    f = sh.formula(r, cc)
                    if f and ("SUBTOTAL(" in f.upper() or "AGGREGATE(" in f.upper()):
                        continue
                yield self.cell_value(sh.name, r, cc, ctx.depth)

    def _scalar(self, v, ctx: _Ctx):
        if isinstance(v, Area):
            if not v.is_single():
                raise NotEvaluable("range used as a scalar (implicit intersection)")
            sh = self.book.sheet(v.sheet)
            c = (ctx.remap or {}).get((sh.name if sh else v.sheet, v.c1), v.c1)
            return self.cell_value(v.sheet, v.r1, c, ctx.depth)
        return v

    # -- evaluation --------------------------------------------------------
    def evaluate(self, text: str, sheet: str, row: int, col: int, remap: Optional[dict] = None):
        """Return ("ok", value) or ("na", reason). value is float|str|bool|None|XLError."""
        try:
            ast = self.parse(text)
            res = self._eval(ast, _Ctx(sheet, row, col, remap, 0))
            res = self._scalar(res, _Ctx(sheet, row, col, remap, 0))
            if isinstance(res, float) and (math.isnan(res) or math.isinf(res)):
                return "ok", XLError("#NUM!")
            return "ok", res
        except NotEvaluable as exc:
            return "na", str(exc)
        except RecursionError:
            return "na", "recursion limit"
        except (ZeroDivisionError, OverflowError, ValueError, TypeError) as exc:
            return "na", f"evaluation error: {exc}"

    def _eval(self, node, ctx: _Ctx):
        k = node[0]
        if k == "num":
            return node[1]
        if k == "str":
            return node[1]
        if k == "bool":
            return node[1]
        if k == "err":
            return XLError(node[1])
        if k in ("ref", "struct"):
            return self.resolve(node, ctx.sheet, ctx.row, ctx.col)
        if k == "name":
            raise NotEvaluable(f"defined name {node[1]!r}")
        if k == "missing":
            return None
        if k == "neg" or k == "pos" or k == "pct":
            v = _to_number(self._scalar(self._eval(node[1], ctx), ctx))
            if isinstance(v, XLError):
                return v
            return -v if k == "neg" else (v / 100.0 if k == "pct" else v)
        if k == "bin":
            return self._binop(node[1], node[2], node[3], ctx)
        if k == "func":
            return self._func(node[1], node[2], ctx)
        raise NotEvaluable(f"unknown node {k}")

    def _binop(self, op, a_node, b_node, ctx):
        a = self._scalar(self._eval(a_node, ctx), ctx)
        b = self._scalar(self._eval(b_node, ctx), ctx)
        if op == "&":
            for v in (a, b):
                if isinstance(v, XLError):
                    return v
            return _to_text(a) + _to_text(b)
        if op in ("=", "<>", "<", ">", "<=", ">="):
            for v in (a, b):
                if isinstance(v, XLError):
                    return v
            return _compare(op, a, b)
        x, y = _to_number(a), _to_number(b)
        if isinstance(x, XLError):
            return x
        if isinstance(y, XLError):
            return y
        if op == "+":
            return x + y
        if op == "-":
            return x - y
        if op == "*":
            return x * y
        if op == "/":
            return XLError("#DIV/0!") if y == 0 else x / y
        if op == "^":
            try:
                return float(x ** y)
            except (OverflowError, ZeroDivisionError):
                return XLError("#NUM!")
        raise NotEvaluable(f"operator {op}")

    def _numbers_from_args(self, args, ctx, skip_hidden=False, skip_subtotals=False, count_mode=None):
        """Collect values like Excel aggregate functions do.

        count_mode None -> list of numbers (errors propagate);
        'counta' -> count of non-empty; 'count' -> count of numbers.
        """
        nums: List[float] = []
        counta = 0
        for a in args:
            if a[0] == "missing":
                continue
            v = self._eval(a, ctx)
            if isinstance(v, Area):
                for cv in self._area_cells(v, ctx, skip_hidden, skip_subtotals):
                    if cv is None:
                        continue
                    counta += 1
                    if isinstance(cv, XLError):
                        if count_mode is None:
                            return cv, 0
                        continue
                    if _is_num(cv):
                        nums.append(float(cv))
                    elif isinstance(cv, (_dt.date, _dt.datetime)):
                        if count_mode is None:
                            raise NotEvaluable("date inside aggregate range")
            else:
                if v is None:
                    continue
                counta += 1
                if isinstance(v, XLError):
                    if count_mode is None:
                        return v, 0
                    continue
                n = _to_number(v)
                if isinstance(n, XLError):
                    if count_mode is None:
                        return n, 0
                    continue
                nums.append(n)
        if count_mode == "counta":
            return float(counta), counta
        if count_mode == "count":
            return float(len(nums)), counta
        return nums, counta

    def _agg(self, kind, args, ctx, skip_hidden=False, skip_subtotals=False):
        if kind in ("counta", "count"):
            v, _ = self._numbers_from_args(args, ctx, skip_hidden, skip_subtotals, count_mode=kind)
            return v
        nums, _ = self._numbers_from_args(args, ctx, skip_hidden, skip_subtotals)
        if isinstance(nums, XLError):
            return nums
        if kind == "sum":
            return float(sum(nums))
        if kind == "average":
            return XLError("#DIV/0!") if not nums else sum(nums) / len(nums)
        if kind == "min":
            return min(nums) if nums else 0.0
        if kind == "max":
            return max(nums) if nums else 0.0
        if kind == "product":
            p = 1.0
            for x in nums:
                p *= x
            return p if nums else 0.0
        raise NotEvaluable(kind)

    def _func(self, name, args, ctx):
        simple = {"SUM": "sum", "COUNTA": "counta", "COUNT": "count", "AVERAGE": "average",
                  "MIN": "min", "MAX": "max", "PRODUCT": "product"}
        if name in simple:
            return self._agg(simple[name], args, ctx)
        if name == "SUBTOTAL":
            if not args:
                raise NotEvaluable("SUBTOTAL without arguments")
            fnv = _to_number(self._scalar(self._eval(args[0], ctx), ctx))
            if isinstance(fnv, XLError):
                return fnv
            fn = int(fnv)
            kind = _SUBTOTAL_KIND.get(fn % 100)
            if kind is None or fn not in (1, 2, 3, 4, 5, 9, 101, 102, 103, 104, 105, 109):
                raise NotEvaluable(f"SUBTOTAL({fn})")
            sh = self.book.sheet(ctx.sheet)
            skip_hidden = fn >= 100 or bool(sh is not None and sh.has_filter)
            return self._agg(kind, args[1:], ctx, skip_hidden=skip_hidden, skip_subtotals=True)
        if name in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
            if len(args) not in (1, 2):
                raise NotEvaluable("ROUND arity")
            x = _to_number(self._scalar(self._eval(args[0], ctx), ctx))
            d = _to_number(self._scalar(self._eval(args[1], ctx), ctx)) if len(args) == 2 else 0.0
            for v in (x, d):
                if isinstance(v, XLError):
                    return v
            q = 10 ** int(d)
            if name == "ROUND":
                return math.floor(abs(x) * q + 0.5) / q * (1 if x >= 0 else -1)
            if name == "ROUNDUP":
                return math.ceil(abs(x) * q) / q * (1 if x >= 0 else -1)
            return math.floor(abs(x) * q) / q * (1 if x >= 0 else -1)
        if name in ("ABS", "INT"):
            if len(args) != 1:
                raise NotEvaluable(f"{name} arity")
            x = _to_number(self._scalar(self._eval(args[0], ctx), ctx))
            if isinstance(x, XLError):
                return x
            return abs(x) if name == "ABS" else float(math.floor(x))
        if name == "IF":
            if len(args) not in (2, 3):
                raise NotEvaluable("IF arity")
            cond = self._scalar(self._eval(args[0], ctx), ctx)
            if isinstance(cond, XLError):
                return cond
            if isinstance(cond, str):
                return XLError("#VALUE!")
            truth = bool(cond) if cond is not None else False
            if truth:
                return self._scalar(self._eval(args[1], ctx), ctx)
            if len(args) == 3:
                return self._scalar(self._eval(args[2], ctx), ctx)
            return False
        if name == "IFERROR":
            if len(args) != 2:
                raise NotEvaluable("IFERROR arity")
            v = self._scalar(self._eval(args[0], ctx), ctx)
            if isinstance(v, XLError):
                return self._scalar(self._eval(args[1], ctx), ctx)
            return v
        if name == "NA" and not args:
            return XLError("#N/A")
        raise NotEvaluable(f"function {name}")


def _compare(op, a, b):
    def rank(v):
        if v is None:
            return (0, 0.0)
        if isinstance(v, bool):
            return (3, float(v))
        if _is_num(v):
            return (1, float(v))
        return (2, str(v).lower())

    if a is None and isinstance(b, str):
        a = ""
    if b is None and isinstance(a, str):
        b = ""
    ra, rb = rank(a), rank(b)
    if ra[0] == 0 and rb[0] == 1:
        ra = (1, 0.0)
    if rb[0] == 0 and ra[0] == 1:
        rb = (1, 0.0)
    return {"=": ra == rb, "<>": ra != rb, "<": ra < rb, ">": ra > rb, "<=": ra <= rb, ">=": ra >= rb}[op]


def values_equal(a: Any, b: Any, tol: float = TOLERANCE) -> bool:
    """Compare an evaluated value with a cached one (numbers within tolerance)."""
    ea, eb = _as_error(a), _as_error(b)
    if ea is not None or eb is not None:
        return ea == eb
    if _is_num(a) and _is_num(b):
        return abs(float(a) - float(b)) <= tol
    if a in (None, "") and b in (None, ""):
        return True
    if _is_num(a) and b is None:
        return abs(float(a)) <= tol
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    return str(a) == str(b)
