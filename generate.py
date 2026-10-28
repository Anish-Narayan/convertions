#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SQL Query Builder 3.0  -  pick and fill in, no SQL knowledge needed
===================================================================
One scrolling page (Tkinter, standard library only) with a live SQL preview.

    1  Target and source   -> schema, INSERT target, FROM table + alias
    2  Joins               -> choose join type, table, alias, then "link" rows:  t1 . id  =  t2 . id
    3  Columns             -> choose table, type column (or "Quick add" many at once)
    4  Filters             -> choose table, type column, CHOOSE an operator, type the value

Example result:
    insert into ${env.schema}.target_table
    select 'abc' as abc
      ,t1.a as t1_a
      ,t2.b as t2_b
    from ${env.schema}.table1 t1
    left join ${env.schema}.table2 t2
      on t1.id = t2.id
      and t1.code = t2.code
    where t1.d in ('X', 'Y')
      and t1.e >= 5

Saved queries and your last session live in SQLQueryBuilder_saved.xlsx, created next
to this program.  Send the program and that file together to share.

Build an .exe (Windows):
    pip install pyinstaller
    pyinstaller --noconsole --onefile --name SQLQueryBuilder sql_query_builder_v3.py
"""
from __future__ import annotations

import os
import re
import sys
import tkinter as tk
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import asdict, dataclass, field
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
from typing import Callable, Dict, List, Optional, Tuple
from xml.sax.saxutils import escape as xml_escape

# =============================================================================
# SECTION 1 - QUERY MODEL, VALIDATION AND SQL GENERATION (no GUI code)
# =============================================================================

APP_NAME = "SQL Query Builder"
APP_VERSION = "3.0"
AUTOSAVE_DELAY_MS = 1500
DEFAULT_SCHEMA = "${env.schema}"

JOIN_LABELS = [("Left join", "left join"), ("Inner join", "inner join"),
               ("Right join", "right join"), ("Full outer join", "full outer join"),
               ("Cross join", "cross join")]
JOIN_LABEL = {sql: label for label, sql in JOIN_LABELS}
JOIN_SQL = {label: sql for label, sql in JOIN_LABELS}
JOIN_TYPES = [sql for _l, sql in JOIN_LABELS]
CROSS_JOIN = "cross join"

# (text shown in the list, SQL operator)
FRIENDLY_OPS = [
    ("=   equals", "="), ("<>  not equal to", "<>"), (">   greater than", ">"),
    (">=  at least", ">="), ("<   less than", "<"), ("<=  at most", "<="),
    ("IN  one of (a, b, c)", "IN"), ("NOT IN  none of (a, b, c)", "NOT IN"),
    ("LIKE  pattern (use %)", "LIKE"), ("NOT LIKE  not matching pattern", "NOT LIKE"),
    ("BETWEEN  two values", "BETWEEN"), ("NOT BETWEEN  outside two values", "NOT BETWEEN"),
    ("IS NULL  is empty", "IS NULL"), ("IS NOT NULL  has a value", "IS NOT NULL"),
]
OP_LABEL = {sql: label for label, sql in FRIENDLY_OPS}
OP_SQL = {label: sql for label, sql in FRIENDLY_OPS}
NULL_OPERATORS = ("IS NULL", "IS NOT NULL")
BETWEEN_OPERATORS = ("BETWEEN", "NOT BETWEEN")
LIST_OPERATORS = ("IN", "NOT IN")
LINK_OPERATORS = ["=", "<>", "<", ">", "<=", ">="]
CONNECTORS = ("AND", "OR")

VT_AUTO = "Auto"
VT_TEXT = "Text"
VT_NUMBER = "Number"
VT_EXPR = "Column / SQL"
VALUE_TYPES = [VT_AUTO, VT_TEXT, VT_NUMBER, VT_EXPR]

KIND_COLUMN = "column"
KIND_TEXT = "text"
KIND_NUMBER = "number"
KIND_EXPR = "expr"
KIND_LABELS = {KIND_COLUMN: "Column", KIND_TEXT: "Text constant",
               KIND_NUMBER: "Number constant", KIND_EXPR: "SQL expression"}
KIND_BY_LABEL = {label: kind for kind, label in KIND_LABELS.items()}

RESERVED_WORDS = {
    "select", "from", "where", "join", "inner", "left", "right", "full", "outer",
    "cross", "on", "and", "or", "not", "in", "as", "group", "order", "by", "having",
    "union", "insert", "into", "values", "like", "between", "is", "null", "distinct",
    "limit",
}
RAW_WORDS = {"null", "true", "false", "current_date", "current_timestamp", "sysdate"}

_PART = r"(?:`[^`]+`|(?:\$\{[^{}\s]+\}|[\w$])+)"
TABLE_RE = re.compile(rf"^{_PART}(?:\.{_PART})*$")
IDENT_RE = re.compile(r"^(?:[A-Za-z_]\w*|`[^`]+`)$")
ALIAS_RE = re.compile(r"^[A-Za-z_]\w*$")
NUMBER_RE = re.compile(r"^[+-]?(?:(?:0|[1-9]\d*)(?:\.\d+)?|\.\d+)$")
VARIABLE_RE = re.compile(r"^\$\{[^{}\s]+\}$")


@dataclass
class Cond:                      # one filter row
    logic: str = "AND"
    open_b: str = ""
    alias: str = ""
    column: str = ""
    op: str = "="
    vtype: str = VT_AUTO
    value: str = ""
    value2: str = ""
    close_b: str = ""


@dataclass
class Link:                      # one "t1.id = t2.id" row of a join
    left_alias: str = ""
    left_col: str = ""
    op: str = "="
    right_alias: str = ""
    right_col: str = ""


@dataclass
class Join:
    jtype: str = "left join"
    table: str = ""
    alias: str = ""
    links: List[Link] = field(default_factory=list)


@dataclass
class Col:                       # one selected column / constant
    kind: str = KIND_COLUMN
    table_alias: str = ""
    value: str = ""
    out_alias: str = ""


@dataclass
class QueryConfig:
    schema: str = DEFAULT_SCHEMA
    use_insert: bool = True
    target_table: str = ""
    from_table: str = ""
    from_alias: str = "t1"
    joins: List[Join] = field(default_factory=list)
    select_items: List[Col] = field(default_factory=list)
    filters: List[Cond] = field(default_factory=list)
    distinct: bool = False
    uppercase: bool = False
    semicolon: bool = False
    prefix_names: bool = True    # t1.a -> t1_a (off: t1.a -> a)


@dataclass
class Issue:
    level: str
    message: str


@dataclass
class BuildResult:
    sql: Optional[str]
    issues: List[Issue]

    @property
    def errors(self) -> List[Issue]:
        return [i for i in self.issues if i.level == "error"]

    @property
    def warnings(self) -> List[Issue]:
        return [i for i in self.issues if i.level == "warning"]


EXAMPLE = QueryConfig(
    target_table="target_table", from_table="table1", from_alias="t1",
    joins=[Join("left join", "table2", "t2",
                [Link("t1", "id", "=", "t2", "id"), Link("t1", "code", "=", "t2", "code")])],
    select_items=[Col(KIND_TEXT, "", "abc", "abc"), Col(KIND_COLUMN, "t1", "a", ""),
                  Col(KIND_COLUMN, "t2", "b", "")],
    filters=[Cond("AND", "", "t1", "d", "IN", VT_AUTO, "X, Y"),
             Cond("AND", "", "t1", "e", ">=", VT_AUTO, "5")])


def sentence(message: str) -> str:
    message = message.strip()
    if not message:
        return message
    message = message[0].upper() + message[1:]
    return message if message.endswith((".", "!", "?")) else message + "."


def make_kw(uppercase: bool) -> Callable[[str], str]:
    return str.upper if uppercase else str.lower


# ----------------------------------------------------------------------- values
def split_top_commas(text: str) -> List[str]:
    """Split on commas that are not inside 'quotes' or (brackets)."""
    parts: List[str] = []
    buf: List[str] = []
    depth, quote = 0, False
    for ch in text:
        if quote:
            buf.append(ch)
            if ch == "'":
                quote = False
            continue
        if ch == "'":
            quote = True
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "," and depth <= 0:
            parts.append("".join(buf).strip())
            buf = []
            continue
        buf.append(ch)
    parts.append("".join(buf).strip())
    return parts


def is_wrapped(s: str) -> bool:
    """True when the whole text is one bracket group: (....)."""
    if not (s.startswith("(") and s.endswith(")")):
        return False
    depth, quote = 0, False
    for i, ch in enumerate(s):
        if quote:
            if ch == "'":
                quote = False
            continue
        if ch == "'":
            quote = True
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0 and i < len(s) - 1:
                return False
    return depth == 0


def quote_text(token: str) -> str:
    t = token.strip()
    if len(t) >= 2 and t[0] == "'" and t[-1] == "'":
        return t                                   # already quoted: as typed
    return "'" + t.replace("'", "''") + "'"


def fmt_value(token: str, aliases: set) -> str:
    """'Auto': numbers, other.column, functions stay as typed; everything else is quoted."""
    t = token.strip()
    if t[0] == "'":
        return quote_text(t)
    if NUMBER_RE.match(t) or t.lower() in RAW_WORDS or "(" in t:
        return t
    m = re.match(r"^([A-Za-z_]\w*)\.(\w+)$", t)
    if m and m.group(1).lower() in aliases:
        return t
    return quote_text(t)


def fmt_operand(vtype: str, raw: str, aliases: set, what: str = "value") -> str:
    t = raw.strip()
    if not t:
        raise ValueError(f"the {what} is required")
    if vtype == VT_TEXT:
        return quote_text(t)
    if vtype == VT_NUMBER:
        if NUMBER_RE.match(t) or VARIABLE_RE.match(t):
            return t
        raise ValueError(f"'{t}' is not a valid number (set Value type to Text if it is text)")
    if vtype == VT_EXPR:
        return t
    return fmt_value(t, aliases)


def fmt_list(vtype: str, raw: str, aliases: set) -> str:
    t = raw.strip()
    if not t:
        raise ValueError("the list of values is required")
    if is_wrapped(t):
        inner = t[1:-1].strip()
        if vtype == VT_EXPR or re.match(r"select\b", inner, re.I):
            return inner                               # sub-query / expression: untouched
        t = inner
    if vtype == VT_EXPR:
        return t
    items = split_top_commas(t)
    if not all(items):
        raise ValueError("the list has an empty item (check the commas)")
    return ", ".join(fmt_operand(vtype, item, aliases) for item in items)


# ---------------------------------------------------------------- building pieces
def select_item_parts(item: Col, prefix_names: bool) -> Tuple[str, str]:
    """Return (expression, output name) for a selected column or raise ValueError."""
    kind = item.kind
    out = item.out_alias.strip()
    value = item.value.strip()
    if kind == KIND_COLUMN:
        alias = item.table_alias.strip()
        if not alias:
            raise ValueError("choose the table")
        if not value:
            raise ValueError("enter the column name")
        if value == "*":
            if out:
                raise ValueError("'*' (all columns) cannot have an output name")
            return f"{alias}.*", ""
        if not IDENT_RE.match(value):
            raise ValueError(f"'{value}' is not a valid column name")
        expr = f"{alias}.{value}"
        if not out:
            bare = value.strip("`")
            out = re.sub(r"\W+", "_", f"{alias}_{bare}" if prefix_names else bare)
    elif kind == KIND_TEXT:
        if not value:
            raise ValueError("enter the text value")
        expr = quote_text(value)
    elif kind == KIND_NUMBER:
        if not (NUMBER_RE.match(value) or VARIABLE_RE.match(value)):
            raise ValueError(f"'{value}' is not a valid number")
        expr = value
    elif kind == KIND_EXPR:
        if not value:
            raise ValueError("enter the SQL expression")
        expr = value
    else:
        raise ValueError(f"unknown column type '{kind}'")
    if not out:
        raise ValueError("enter an output name for this kind of column")
    if not IDENT_RE.match(out):
        raise ValueError(f"'{out}' is not a valid output name")
    if not out.startswith("`") and out.lower() in RESERVED_WORDS:
        raise ValueError(f"'{out}' is a reserved SQL word - choose another output name")
    return expr, out


def render_cond(c: Cond, aliases: set, kw: Callable[[str], str]) -> str:
    """One filter row -> SQL text (no connector, no brackets)."""
    column, alias = c.column.strip(), c.alias.strip()
    if not column:
        raise ValueError("choose or type the column")
    if alias:
        if not ALIAS_RE.match(alias):
            raise ValueError(f"'{alias}' is not a valid table alias")
        if alias.lower() not in aliases:
            raise ValueError(f"table '{alias}' is not defined in step 1 or 2")
        if not IDENT_RE.match(column):
            raise ValueError(f"'{column}' is not a valid column name "
                             "(leave Table empty if you want to type an expression)")
        left = f"{alias}.{column}"
    else:
        left = column
    op = c.op
    if op not in OP_LABEL:
        raise ValueError(f"unknown operator '{op}'")
    if op in NULL_OPERATORS:
        return f"{left} {kw(op)}"
    if op in BETWEEN_OPERATORS:
        low = fmt_operand(c.vtype, c.value, aliases, "first value")
        high = fmt_operand(c.vtype, c.value2, aliases, "second value")
        return f"{left} {kw(op)} {low} {kw('and')} {high}"
    if op in LIST_OPERATORS:
        return f"{left} {kw(op)} ({fmt_list(c.vtype, c.value, aliases)})"
    return f"{left} {kw(op)} {fmt_operand(c.vtype, c.value, aliases)}"


def render_link(link: Link, allowed: set, everything: set) -> str:
    """One join link row -> 't1.id = t2.id'."""
    parts = []
    for side, alias, col in (("left", link.left_alias, link.left_col),
                             ("right", link.right_alias, link.right_col)):
        alias, col = alias.strip(), col.strip()
        if not alias:
            raise ValueError(f"choose the {side} table")
        if not col:
            raise ValueError(f"enter the {side} column")
        if not IDENT_RE.match(col):
            raise ValueError(f"'{col}' is not a valid column name")
        if alias.lower() not in everything:
            raise ValueError(f"table '{alias}' is not defined in step 1 or 2")
        if alias.lower() not in allowed:
            raise ValueError(f"table '{alias}' belongs to a later join and cannot be used here")
        parts.append(f"{alias}.{col}")
    if link.op not in LINK_OPERATORS:
        raise ValueError(f"unknown operator '{link.op}'")
    return f"{parts[0]} {link.op} {parts[1]}"


def cond_is_blank(c: Cond) -> bool:
    return not (c.column.strip() or c.value.strip() or c.value2.strip())


def link_is_blank(l: Link) -> bool:
    return not (l.left_col.strip() or l.right_col.strip())


def col_is_blank(c: Col) -> bool:
    return not (c.value.strip() or c.out_alias.strip())


# --------------------------------------------------------------------- build query
def build_query(cfg: QueryConfig) -> BuildResult:
    issues: List[Issue] = []

    def err(msg: str) -> None:
        issues.append(Issue("error", sentence(msg)))

    def warn(msg: str) -> None:
        issues.append(Issue("warning", sentence(msg)))

    kw = make_kw(cfg.uppercase)

    schema = cfg.schema.strip()
    if schema and not TABLE_RE.match(schema):
        err(f"Schema / prefix '{schema}' is not valid (box 1)")

    def qualify(name: str) -> str:
        outside = re.sub(r"`[^`]*`|\$\{[^{}]*\}", "x", name)
        return name if (not schema or "." in outside) else f"{schema}.{name}"

    def check_alias(alias: str, where: str) -> None:
        alias = alias.strip()
        if not alias:
            err(f"{where}: alias is required")
        elif not ALIAS_RE.match(alias):
            err(f"{where}: alias '{alias}' is not valid (letters, digits and _ only; "
                "must not start with a digit)")
        elif alias.lower() in RESERVED_WORDS:
            err(f"{where}: '{alias}' is a reserved SQL word and cannot be an alias")

    # ---- target
    target_sql = ""
    if cfg.use_insert:
        target = cfg.target_table.strip()
        if not target:
            err("Target table is required (box 1) - or untick INSERT INTO")
        elif not TABLE_RE.match(target):
            err(f"Target table '{target}' is not a valid table name (box 1)")
        else:
            target_sql = qualify(target)

    # ---- FROM
    from_table = cfg.from_table.strip()
    from_sql = ""
    if not from_table:
        err("FROM table is required (box 1)")
    elif not TABLE_RE.match(from_table):
        err(f"FROM table '{from_table}' is not a valid table name (box 1)")
    else:
        from_sql = qualify(from_table)
    from_alias = cfg.from_alias.strip()
    check_alias(from_alias, "FROM table")

    # ---- aliases
    ordered = [from_alias.lower()]
    for i, join in enumerate(cfg.joins, 1):
        check_alias(join.alias, f"Join {i}")
        ordered.append(join.alias.strip().lower())
    for alias in sorted({a for a in ordered if a}):
        if ordered.count(alias) > 1:
            err(f"Table alias '{alias}' is used {ordered.count(alias)} times - every table needs its own alias")
    everything = {a for a in ordered if a}

    lines: List[str] = []
    if target_sql:
        lines.append(f"{kw('insert into')} {target_sql}")

    # ---- SELECT
    select_lines: List[str] = []
    out_names: Dict[str, List[int]] = {}
    for n, item in enumerate(cfg.select_items, 1):
        if col_is_blank(item):
            continue
        try:
            expr, out = select_item_parts(item, cfg.prefix_names)
        except ValueError as exc:
            err(f"Column {n}: {exc}")
            continue
        if item.kind == KIND_COLUMN and item.table_alias.strip().lower() not in everything:
            err(f"Column {n}: table '{item.table_alias.strip()}' is not defined in step 1 or 2")
            continue
        select_lines.append(f"{expr} {kw('as')} {out}" if out else expr)
        if out:
            out_names.setdefault(out.lower(), []).append(n)
    if not any(not col_is_blank(c) for c in cfg.select_items):
        err("Add at least one column to select (step 3)")
    for name, positions in out_names.items():
        if len(positions) > 1:
            warn(f"Output name '{name}' is used by columns {', '.join(map(str, positions))}. "
                 "Duplicate names can fail in some databases")
    select_word = kw("select") + (f" {kw('distinct')}" if cfg.distinct else "")
    for idx, text in enumerate(select_lines):
        lines.append(f"{select_word} {text}" if idx == 0 else f"  ,{text}")

    # ---- FROM + JOINs
    if from_sql:
        lines.append(f"{kw('from')} {from_sql} {from_alias}")
    for i, join in enumerate(cfg.joins, 1):
        label = f"Join {i} ({join.alias.strip() or '?'})"
        if join.jtype not in JOIN_TYPES:
            err(f"{label}: unknown join type '{join.jtype}'")
            continue
        table = join.table.strip()
        if not table:
            err(f"{label}: enter the table name")
            continue
        if not TABLE_RE.match(table):
            err(f"{label}: '{table}' is not a valid table name")
            continue
        lines.append(f"{kw(join.jtype)} {qualify(table)} {join.alias.strip()}")
        used_links = [(m, l) for m, l in enumerate(join.links, 1) if not link_is_blank(l)]
        if join.jtype == CROSS_JOIN:
            if used_links:
                warn(f"{label}: link rows are ignored for a cross join")
            continue
        if not used_links:
            err(f"{label}: add at least one link, for example {from_alias or 't1'}.id = "
                f"{join.alias.strip() or 't2'}.id")
            continue
        allowed = {a for a in ordered[: i + 1] if a}
        on_lines: List[str] = []
        for m, link in used_links:
            try:
                text = render_link(link, allowed, everything)
            except ValueError as exc:
                err(f"{label}, link {m}: {exc}")
                continue
            on_lines.append(f"  {kw('on') if not on_lines else kw('and')} {text}")
        lines.extend(on_lines)

    # ---- WHERE
    depth, first = 0, True
    for n, cond in enumerate(cfg.filters, 1):
        if cond_is_blank(cond):
            continue
        where = f"Filter {n}"
        opening, closing = cond.open_b.strip(), cond.close_b.strip()
        if opening.strip("(") or closing.strip(")"):
            err(f"{where}: brackets may only contain ( or ) characters")
            opening = closing = ""
        if not first and cond.logic not in CONNECTORS:
            err(f"{where}: choose AND or OR")
            continue
        try:
            text = render_cond(cond, everything, kw)
        except ValueError as exc:
            err(f"{where}: {exc}")
            continue
        depth += len(opening) - len(closing)
        if depth < 0:
            err(f"{where}: a closing bracket has no matching opening bracket")
            depth = 0
        body = f"{opening}{text}{closing}"
        lines.append(f"{kw('where')} {body}" if first else f"  {kw(cond.logic)} {body}")
        first = False
    if depth:
        err("Filters: brackets are not balanced - every ( needs a matching )")

    if any(i.level == "error" for i in issues):
        return BuildResult(None, issues)
    sql = "\n".join(lines)
    return BuildResult(sql + (";" if cfg.semicolon else ""), issues)


def config_to_dict(cfg: QueryConfig) -> dict:
    return asdict(cfg)


# =============================================================================
# EXCEL STORAGE (standard library only - an .xlsx file is a zip of XML files)
# =============================================================================
DATA_FILE_NAME = "SQLQueryBuilder_saved.xlsx"
SESSION_NAME = "(Last session)"
MAX_QUERY_NAME_LENGTH = 100
MAX_CELL_LENGTH = 30000

SHEET_README = "Read me"
SHEET_QUERIES = "Queries"
QUERIES_HEADER = ["Name", "Saved", "Schema / prefix", "Generate INSERT", "Target table",
                  "From table", "From alias", "Joins", "Columns", "Filters",
                  "SELECT DISTINCT", "UPPERCASE keywords", "End with ;",
                  "Prefix column names with alias", "Generated SQL (information only)"]
QUERIES_WIDTHS = [28, 17, 18, 16, 22, 22, 10, 45, 40, 45, 16, 20, 11, 28, 70]
README_LINES = [
    "This workbook is created and updated by the SQL Query Builder program.",
    "Keep it in the same folder as the program (.py or .exe) and send both files together.",
    "",
    "Queries: one row per saved query. The row '(Last session)' is restored when the program starts.",
    "Joins / Columns / Filters hold the rows of the program, one row per line, fields separated by tabs.",
    "The last column shows the SQL that was generated (information only - it is not read back).",
    "",
    "Please change queries in the program, not here. Do not rename the sheets or the header row.",
    "Close this file in Excel before using the program, otherwise the program cannot save.",
]

_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_XML_ILLEGAL = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


class StoreError(Exception):
    """The Excel data file exists but cannot be understood."""


def app_folder() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def describe_store_error(exc: Exception) -> str:
    if isinstance(exc, PermissionError):
        return f"{exc}\n\nIf the Excel data file is open in Excel, close it and try again."
    return str(exc)


def _col_letter(index: int) -> str:
    letters, index = "", index + 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _cell_xml(ref: str, value, style: int = 0) -> str:
    if value is None or value == "":
        return ""
    attr = f' s="{style}"' if style else ""
    if isinstance(value, int) and not isinstance(value, bool):
        return f'<c r="{ref}"{attr}><v>{value}</v></c>'
    text = xml_escape(_XML_ILLEGAL.sub("", str(value))[:MAX_CELL_LENGTH])
    return f'<c r="{ref}"{attr} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def _sheet_xml(header: List[str], rows: List[list], widths: List[int], wrap: bool) -> str:
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
           '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
           'activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/>'
           '</sheetView></sheetViews>',
           "<cols>" + "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
                              for i, w in enumerate(widths, 1)) + "</cols>",
           "<sheetData>",
           '<row r="1">' + "".join(_cell_xml(f"{_col_letter(i)}1", h, 1)
                                   for i, h in enumerate(header)) + "</row>"]
    body_style = 2 if wrap else 0
    for r, row in enumerate(rows, 2):
        cells = "".join(_cell_xml(f"{_col_letter(i)}{r}", v, body_style) for i, v in enumerate(row))
        out.append(f'<row r="{r}">{cells}</row>' if cells else f'<row r="{r}"/>')
    out.append("</sheetData></worksheet>")
    return "".join(out)


_STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="2"><font><sz val="10"/><name val="Arial"/></font>'
    '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts>'
    '<fills count="3"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3A5F"/><bgColor indexed="64"/>'
    '</patternFill></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">'
    '<alignment vertical="top" wrapText="1"/></xf></cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>'
)


def write_xlsx(path: str, sheets: List[Tuple[str, List[str], List[list], List[int], bool]]) -> None:
    """Write [(name, header, rows, widths, wrap)] atomically (temp file, then move into place)."""
    n = len(sheets)
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                  for i in range(1, n + 1))
        + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>')
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'xmlns:r="{rel}"><sheets>'
        + "".join(f'<sheet name="{xml_escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
                  for i, (name, *_rest) in enumerate(sheets, 1))
        + '</sheets></workbook>')
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(f'<Relationship Id="rId{i}" Type="{rel}/worksheet" Target="worksheets/sheet{i}.xml"/>'
                  for i in range(1, n + 1))
        + f'<Relationship Id="rId{n + 1}" Type="{rel}/styles" Target="styles.xml"/></Relationships>')

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = path + ".tmp"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", content_types)
            archive.writestr("_rels/.rels", root_rels)
            archive.writestr("xl/workbook.xml", workbook)
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
            archive.writestr("xl/styles.xml", _STYLES_XML)
            for i, (_n, header, rows, widths, wrap) in enumerate(sheets, 1):
                archive.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(header, rows, widths, wrap))
        os.replace(tmp, path)
    except BaseException:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def _rich_text(node) -> str:
    parts: List[str] = []
    for child in node:
        if child.tag == _MAIN_NS + "t":
            parts.append(child.text or "")
        elif child.tag == _MAIN_NS + "r":
            parts.extend(t.text or "" for t in child.findall(_MAIN_NS + "t"))
    return "".join(parts)


def _plain_number(raw: str) -> str:
    try:
        number = float(raw)
    except ValueError:
        return raw
    return str(int(number)) if number.is_integer() and abs(number) < 1e15 else repr(number)


def _read_sheet(root, shared: List[str]) -> List[List[str]]:
    rows: List[List[str]] = []
    for row in root.iter(_MAIN_NS + "row"):
        values: Dict[int, str] = {}
        for cell in row.findall(_MAIN_NS + "c"):
            match = re.match(r"([A-Z]+)", cell.get("r") or "")
            if not match:
                continue
            kind = cell.get("t")
            if kind == "inlineStr":
                node = cell.find(_MAIN_NS + "is")
                text = _rich_text(node) if node is not None else ""
            else:
                v = cell.find(_MAIN_NS + "v")
                raw = v.text if v is not None and v.text is not None else ""
                if kind == "s":
                    try:
                        text = shared[int(raw)]
                    except (ValueError, IndexError):
                        text = ""
                elif kind == "str":
                    text = raw
                elif kind == "b":
                    text = "TRUE" if raw == "1" else "FALSE"
                elif kind == "e":
                    text = ""
                else:
                    text = _plain_number(raw)
            values[_col_index(match.group(1))] = text
        if any(v != "" for v in values.values()):
            rows.append([values.get(i, "") for i in range(max(values) + 1)])
    return rows


def read_xlsx(path: str) -> Dict[str, List[List[str]]]:
    try:
        with zipfile.ZipFile(path) as archive:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            targets = {r.get("Id"): r.get("Target") for r in rels.iter(_PKG_NS + "Relationship")}
            shared: List[str] = []
            if "xl/sharedStrings.xml" in archive.namelist():
                root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                shared = [_rich_text(si) for si in root.iter(_MAIN_NS + "si")]
            sheets: Dict[str, List[List[str]]] = {}
            for sheet in workbook.iter(_MAIN_NS + "sheet"):
                name, target = sheet.get("name"), targets.get(sheet.get(_REL_NS + "id"))
                if not name or not target:
                    continue
                part = target.lstrip("/") if target.startswith("/") else "xl/" + target
                sheets[name] = _read_sheet(ET.fromstring(archive.read(part)), shared)
            return sheets
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        raise StoreError(f"'{os.path.basename(path)}' is not a readable Excel workbook ({exc})") from exc


# ---- the rows of a query <-> text cells (one row per line, fields separated by tabs)
def _yes_no(flag: bool) -> str:
    return "Yes" if flag else "No"


def _parse_bool(text: str, default: bool) -> bool:
    t = text.strip().lower()
    if t in ("yes", "y", "true", "1", "x"):
        return True
    if t in ("no", "n", "false", "0"):
        return False
    return default


def _flat(value) -> str:
    return str(value).replace("\t", " ").replace("\r", " ").replace("\n", " ")


def _line(*fields) -> str:
    return "\t".join(_flat(f) for f in fields)


def _fields(line: str, count: int) -> List[str]:
    return (line.split("\t") + [""] * count)[:count]


def _pick(text: str, choices: List[str], default: str = "") -> str:
    cleaned = " ".join(text.split())
    for choice in choices:
        if choice.casefold() == cleaned.casefold():
            return choice
    return cleaned or default


def joins_to_text(joins: List[Join]) -> str:
    out = []
    for j in joins:
        out.append(_line("J", j.jtype, j.table, j.alias))
        out.extend(_line("L", l.left_alias, l.left_col, l.op, l.right_alias, l.right_col) for l in j.links)
    return "\n".join(out)


def joins_from_text(text: str) -> List[Join]:
    joins: List[Join] = []
    for line in text.splitlines():
        f = _fields(line, 6)
        if f[0] == "J":
            joins.append(Join(_pick(f[1], JOIN_TYPES, "left join"), f[2].strip(), f[3].strip(), []))
        elif f[0] == "L" and joins:
            joins[-1].links.append(Link(f[1].strip(), f[2].strip(), _pick(f[3], LINK_OPERATORS, "="),
                                        f[4].strip(), f[5].strip()))
    return joins


def columns_to_text(items: List[Col]) -> str:
    return "\n".join(_line(c.kind, c.table_alias, c.value, c.out_alias) for c in items)


def columns_from_text(text: str) -> List[Col]:
    out = []
    for line in text.splitlines():
        f = _fields(line, 4)
        out.append(Col(_pick(f[0], list(KIND_LABELS), KIND_COLUMN), f[1].strip(), f[2].strip(), f[3].strip()))
    return out


def filters_to_text(items: List[Cond]) -> str:
    return "\n".join(_line(c.logic, c.open_b, c.alias, c.column, c.op, c.vtype, c.value, c.value2, c.close_b)
                     for c in items)


def filters_from_text(text: str) -> List[Cond]:
    out = []
    for line in text.splitlines():
        f = _fields(line, 9)
        out.append(Cond(_pick(f[0], list(CONNECTORS), "AND"), f[1].strip(), f[2].strip(), f[3].strip(),
                        _pick(f[4], list(OP_LABEL), "="), _pick(f[5], VALUE_TYPES, VT_AUTO),
                        f[6].strip(), f[7].strip(), f[8].strip()))
    return out


class QueryStore:
    """Saved queries + the last session, kept in one Excel workbook next to the program."""

    def __init__(self, path: Optional[str] = None):
        self.path = path or os.path.join(app_folder(), DATA_FILE_NAME)

    def _load_all(self) -> Dict[str, Tuple[str, QueryConfig]]:
        if not os.path.exists(self.path):
            return {}
        sheets = read_xlsx(self.path)
        if SHEET_QUERIES not in sheets or not sheets[SHEET_QUERIES]:
            raise StoreError(f"'{os.path.basename(self.path)}' has no '{SHEET_QUERIES}' sheet - "
                             "it was not created by this program. Move it away or rename it.")
        rows = sheets[SHEET_QUERIES]
        header = [h.strip().lower() for h in rows[0]]
        entries: Dict[str, Tuple[str, QueryConfig]] = {}
        for r in rows[1:]:
            row = {header[i]: (r[i] if i < len(r) else "") for i in range(len(header)) if header[i]}

            def cell(title: str) -> str:
                return row.get(title.lower(), "").strip()

            name = cell("Name")
            if not name or name in entries:
                continue
            entries[name] = (cell("Saved"), QueryConfig(
                schema=cell("Schema / prefix"),
                use_insert=_parse_bool(cell("Generate INSERT"), True),
                target_table=cell("Target table"),
                from_table=cell("From table"),
                from_alias=cell("From alias") or "t1",
                joins=joins_from_text(cell("Joins")),
                select_items=columns_from_text(cell("Columns")),
                filters=filters_from_text(cell("Filters")),
                distinct=_parse_bool(cell("SELECT DISTINCT"), False),
                uppercase=_parse_bool(cell("UPPERCASE keywords"), False),
                semicolon=_parse_bool(cell("End with ;"), False),
                prefix_names=_parse_bool(cell("Prefix column names with alias"), True)))
        return entries

    def _write_all(self, entries: Dict[str, Tuple[str, QueryConfig]]) -> None:
        rows = []
        for name, (saved, c) in entries.items():
            sql = build_query(c).sql
            rows.append([name, saved, c.schema, _yes_no(c.use_insert), c.target_table, c.from_table,
                         c.from_alias, joins_to_text(c.joins), columns_to_text(c.select_items),
                         filters_to_text(c.filters), _yes_no(c.distinct), _yes_no(c.uppercase),
                         _yes_no(c.semicolon), _yes_no(c.prefix_names),
                         sql if sql is not None else "(query not complete)"])
        write_xlsx(self.path, [
            (SHEET_README, ["SQL Query Builder - data file"], [[ln] for ln in README_LINES], [110], False),
            (SHEET_QUERIES, QUERIES_HEADER, rows, QUERIES_WIDTHS, True),
        ])

    @staticmethod
    def _now() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M")

    @staticmethod
    def clean_name(name: str) -> str:
        name = " ".join(name.split())
        if not name:
            raise ValueError("enter a name for the query")
        if len(name) > MAX_QUERY_NAME_LENGTH:
            raise ValueError(f"the name is too long (maximum {MAX_QUERY_NAME_LENGTH} characters)")
        if name.casefold() == SESSION_NAME.casefold():
            raise ValueError(f"the name {SESSION_NAME} is reserved - choose another name")
        return name

    def save_session(self, cfg: QueryConfig) -> None:
        entries = self._load_all()
        entries[SESSION_NAME] = (self._now(), cfg)
        self._write_all(entries)

    def load_session(self) -> Optional[QueryConfig]:
        entry = self._load_all().get(SESSION_NAME)
        return entry[1] if entry else None

    def list_queries(self) -> List[str]:
        return sorted((n for n in self._load_all() if n != SESSION_NAME), key=str.lower)

    def has_query(self, name: str) -> bool:
        return self.clean_name(name) in self._load_all()

    def save_query(self, name: str, cfg: QueryConfig) -> str:
        name = self.clean_name(name)
        entries = self._load_all()
        entries[name] = (self._now(), cfg)
        self._write_all(entries)
        return name

    def load_query(self, name: str) -> QueryConfig:
        entry = self._load_all().get(name)
        if entry is None or name == SESSION_NAME:
            raise ValueError(f"no saved query named '{name}'")
        return entry[1]

    def delete_query(self, name: str) -> None:
        entries = self._load_all()
        if name != SESSION_NAME and entries.pop(name, None) is not None:
            self._write_all(entries)


# =============================================================================
# SECTION 2 - GUI: THEME, SCROLLING PAGE, ROW EDITORS
# =============================================================================

class Theme:
    PRIMARY = "#1F3A5F"
    ACCENT = "#2F6FDE"
    ACCENT_DARK = "#2358B5"
    BG = "#F3F5F9"
    CARD = "#FFFFFF"
    BORDER = "#C9D1DC"
    TEXT = "#1B2430"
    MUTED = "#5B6573"
    ERROR = "#B42318"
    WARNING = "#B54708"
    OK = "#067647"
    EDITOR_BG = "#18202C"
    EDITOR_FG = "#E6EDF3"
    EDITOR_MUTED = "#8B98A9"
    EDITOR_KEYWORD = "#79B8FF"
    EDITOR_STRING = "#9AE6B4"
    EDITOR_VARIABLE = "#FFB86C"


def enable_high_dpi() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def configure_style(root: tk.Misc) -> None:
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        try:
            tkfont.nametofont(name).configure(size=10)
        except tk.TclError:
            pass
    default = tkfont.nametofont("TkDefaultFont")
    families = set(tkfont.families(root))
    mono = next((f for f in ("Consolas", "Cascadia Mono", "Menlo", "DejaVu Sans Mono",
                             "Courier New") if f in families), None)
    fixed = tkfont.nametofont("TkFixedFont")
    if mono:
        fixed.configure(family=mono)
    fixed.configure(size=11)
    bold = default.copy()
    bold.configure(weight="bold")
    title = default.copy()
    title.configure(size=17, weight="bold")

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    T = Theme
    style.configure(".", background=T.BG, foreground=T.TEXT, font=default)
    style.configure("TFrame", background=T.BG)
    style.configure("TLabel", background=T.BG, foreground=T.TEXT)
    style.configure("Muted.TLabel", foreground=T.MUTED)
    style.configure("Header.TFrame", background=T.PRIMARY)
    style.configure("Header.TLabel", background=T.PRIMARY, foreground="white", font=title)
    style.configure("SubHeader.TLabel", background=T.PRIMARY, foreground="#C9D6EA")
    style.configure("Status.TFrame", background="#E4E8EF")
    style.configure("Status.TLabel", background="#E4E8EF", foreground=T.MUTED)
    style.configure("TLabelframe", background=T.BG, bordercolor=T.BORDER)
    style.configure("TLabelframe.Label", background=T.BG, foreground=T.PRIMARY, font=bold)
    style.configure("TCheckbutton", background=T.BG)
    style.configure("TButton", padding=(10, 5))
    style.configure("Tool.TButton", padding=(2, 1))
    style.configure("Accent.TButton", background=T.ACCENT, foreground="white", padding=(14, 6))
    style.map("Accent.TButton",
              background=[("disabled", "#A9B4C6"), ("active", T.ACCENT_DARK)],
              foreground=[("disabled", "#EEF1F6")])
    style.map("TEntry", fieldbackground=[("disabled", "#E9ECF1")])
    style.map("TCombobox", fieldbackground=[("disabled", "#E9ECF1"), ("readonly", T.CARD)])


class ScrollPage(ttk.Frame):
    """A vertically scrolling page; put widgets into ``self.inner``."""

    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, background=Theme.BG)
        bar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=bar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")
        self.inner = ttk.Frame(self.canvas)
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
                        lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfigure(self._window, width=e.width))

    def scroll(self, direction: int) -> None:
        self.canvas.yview_scroll(direction * 2, "units")

    def see(self, widget) -> None:
        """Scroll so that ``widget`` is visible."""
        try:
            self.update_idletasks()
            y, w = 0, widget
            while w is not self.inner:
                y += w.winfo_y()
                w = w.master
            height = widget.winfo_height()
            total = max(1, self.inner.winfo_height())
            top = self.canvas.canvasy(0)
            view = self.canvas.winfo_height()
            if y < top:
                self.canvas.yview_moveto(y / total)
            elif y + height > top + view:
                self.canvas.yview_moveto((y + height - view) / total)
        except (tk.TclError, AttributeError):
            pass


class RowList(ttk.Frame):
    """A list of editable rows with an '+ Add' button.  Rows move / delete themselves."""

    def __init__(self, master, factory: Callable[["RowList"], "RowBase"],
                 on_change: Callable[[], None], add_text: str,
                 see: Optional[Callable] = None):
        super().__init__(master)
        self.factory = factory
        self.on_change = on_change
        self.see = see
        self.rows: List["RowBase"] = []
        self.columnconfigure(0, weight=1)
        self.body = ttk.Frame(self)
        self.body.grid(row=0, column=0, sticky="ew")
        self.body.columnconfigure(0, weight=1)
        self.bar = ttk.Frame(self)
        self.bar.grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Button(self.bar, text=add_text, command=self.add).pack(side="left")

    def add(self, item=None, notify: bool = True) -> "RowBase":
        row = self.factory(self)
        self.rows.append(row)
        if item is None:
            row.set_defaults()
        else:
            row.set_item(item)
        self._regrid()
        if notify:
            self.on_change()
            if self.see:
                self.see(row)
            if row.first_widget is not None:
                row.first_widget.focus_set()
        return row

    def add_after_last(self, row: "RowBase") -> None:
        if self.rows and row is self.rows[-1]:
            self.add()

    def remove(self, row: "RowBase") -> None:
        if row in self.rows:
            self.rows.remove(row)
            row.destroy()
            self._regrid()
            self.on_change()

    def move(self, row: "RowBase", delta: int) -> None:
        i = self.rows.index(row)
        j = i + delta
        if 0 <= j < len(self.rows):
            self.rows[i], self.rows[j] = self.rows[j], self.rows[i]
            self._regrid()
            self.on_change()

    def _regrid(self) -> None:
        for i, row in enumerate(self.rows):
            row.grid(row=i, column=0, sticky="ew", pady=2)
            row.set_index(i)

    def items(self) -> list:
        return [row.get_item() for row in self.rows]

    def set_items(self, items: list) -> None:
        for row in self.rows:
            row.destroy()
        self.rows = []
        for item in items:
            self.add(item, notify=False)
        self.on_change()


class RowBase(ttk.Frame):
    first_widget = None

    def __init__(self, owner: RowList):
        super().__init__(owner.body)
        self.owner = owner

    def set_index(self, index: int) -> None:
        pass

    def set_defaults(self) -> None:
        pass

    def get_item(self):
        raise NotImplementedError

    def set_item(self, item) -> None:
        raise NotImplementedError

    def watch(self, *variables: tk.Variable) -> None:
        for var in variables:
            var.trace_add("write", lambda *_: self.owner.on_change())

    def add_controls(self, parent, column: int) -> None:
        bar = ttk.Frame(parent)
        bar.grid(row=0, column=column, padx=(8, 0))
        for text, command in (("\u25B2", lambda: self.owner.move(self, -1)),
                              ("\u25BC", lambda: self.owner.move(self, 1)),
                              ("\u2715", lambda: self.owner.remove(self))):
            ttk.Button(bar, text=text, width=2, style="Tool.TButton", command=command).pack(
                side="left", padx=1)

    def enter_adds_row(self, widget) -> None:
        widget.bind("<Return>", lambda _e: self.owner.add_after_last(self))


def alias_box(parent, variable: tk.StringVar, app: "QueryBuilderApp", width: int = 5) -> ttk.Combobox:
    box = ttk.Combobox(parent, textvariable=variable, width=width,
                       postcommand=lambda: box.configure(values=app.get_aliases()))
    return box


def column_box(parent, variable: tk.StringVar, alias_var: tk.StringVar, app: "QueryBuilderApp",
               width: int = 14) -> ttk.Combobox:
    box = ttk.Combobox(parent, textvariable=variable, width=width,
                       postcommand=lambda: box.configure(values=app.known_columns(alias_var.get())))
    return box


class CondRow(RowBase):
    """WHERE row:  [AND] [(] table . column  operator  [type] value [value 2] [)]"""

    def __init__(self, owner, app: "QueryBuilderApp"):
        super().__init__(owner)
        self.app = app
        self.v_logic = tk.StringVar(value="AND")
        self.v_open = tk.StringVar()
        self.v_alias = tk.StringVar()
        self.v_col = tk.StringVar()
        self.v_op = tk.StringVar(value=OP_LABEL["="])
        self.v_vt = tk.StringVar(value=VT_AUTO)
        self.v_val = tk.StringVar()
        self.v_val2 = tk.StringVar()
        self.v_close = tk.StringVar()

        pad = {"padx": (0, 4)}
        self.lbl_where = ttk.Label(self, text="WHERE", width=6, style="Muted.TLabel")
        self.cb_logic = ttk.Combobox(self, textvariable=self.v_logic, values=list(CONNECTORS),
                                     state="readonly", width=5)
        self.lbl_where.grid(row=0, column=0, sticky="w", **pad)
        self.cb_logic.grid(row=0, column=0, sticky="w", **pad)
        ttk.Combobox(self, textvariable=self.v_open, values=["", "(", "(("], state="readonly",
                     width=2).grid(row=0, column=1, **pad)
        alias_box(self, self.v_alias, app).grid(row=0, column=2, **pad)
        self.cb_col = column_box(self, self.v_col, self.v_alias, app, 13)
        self.cb_col.grid(row=0, column=3, **pad)
        self.cb_op = ttk.Combobox(self, textvariable=self.v_op, values=[l for l, _s in FRIENDLY_OPS],
                                  state="readonly", width=24)
        self.cb_op.grid(row=0, column=4, **pad)
        self.cb_vt = ttk.Combobox(self, textvariable=self.v_vt, values=VALUE_TYPES, state="readonly",
                                  width=11)
        self.cb_vt.grid(row=0, column=5, **pad)
        self.ent_val = ttk.Entry(self, textvariable=self.v_val, width=14)
        self.ent_val.grid(row=0, column=6, sticky="ew", **pad)
        self.ent_val2 = ttk.Entry(self, textvariable=self.v_val2, width=10)
        self.ent_val2.grid(row=0, column=7, sticky="ew", **pad)
        ttk.Combobox(self, textvariable=self.v_close, values=["", ")", "))"], state="readonly",
                     width=2).grid(row=0, column=8, **pad)
        self.add_controls(self, 9)
        self.columnconfigure(6, weight=1)
        self.first_widget = self.cb_col
        self.enter_adds_row(self.ent_val)
        self.enter_adds_row(self.ent_val2)

        self.watch(self.v_logic, self.v_open, self.v_alias, self.v_col, self.v_vt,
                   self.v_val, self.v_val2, self.v_close)
        self.v_op.trace_add("write", lambda *_: (self._sync(), self.owner.on_change()))
        self._sync()

    def _sync(self) -> None:
        op = OP_SQL.get(self.v_op.get(), "=")
        no_value = op in NULL_OPERATORS
        self.cb_vt.configure(state="disabled" if no_value else "readonly")
        self.ent_val.configure(state="disabled" if no_value else "normal")
        if op in BETWEEN_OPERATORS:
            self.ent_val2.grid()
        else:
            self.ent_val2.grid_remove()

    def set_index(self, index: int) -> None:
        if index == 0:
            self.cb_logic.grid_remove()
            self.lbl_where.grid()
        else:
            self.lbl_where.grid_remove()
            self.cb_logic.grid()

    def set_defaults(self) -> None:
        self.v_alias.set(self.app.default_alias())

    def get_item(self) -> Cond:
        return Cond(self.v_logic.get() or "AND", self.v_open.get().strip(), self.v_alias.get().strip(),
                    self.v_col.get().strip(), OP_SQL.get(self.v_op.get(), "="), self.v_vt.get() or VT_AUTO,
                    self.v_val.get().strip(), self.v_val2.get().strip(), self.v_close.get().strip())

    def set_item(self, c: Cond) -> None:
        self.v_logic.set(c.logic)
        self.v_open.set(c.open_b)
        self.v_alias.set(c.alias)
        self.v_col.set(c.column)
        self.v_op.set(OP_LABEL.get(c.op, OP_LABEL["="]))
        self.v_vt.set(c.vtype)
        self.v_val.set(c.value)
        self.v_val2.set(c.value2)
        self.v_close.set(c.close_b)
        self._sync()


class ColRow(RowBase):
    """Selected column:  [Column|Text|Number|Expression]  table . name   as  output name"""

    def __init__(self, owner, app: "QueryBuilderApp"):
        super().__init__(owner)
        self.app = app
        self.v_kind = tk.StringVar(value=KIND_LABELS[KIND_COLUMN])
        self.v_alias = tk.StringVar()
        self.v_val = tk.StringVar()
        self.v_out = tk.StringVar()
        pad = {"padx": (0, 4)}
        kind = ttk.Combobox(self, textvariable=self.v_kind, values=list(KIND_LABELS.values()),
                            state="readonly", width=14)
        kind.grid(row=0, column=0, **pad)
        self.cb_alias = alias_box(self, self.v_alias, app)
        self.cb_alias.grid(row=0, column=1, **pad)
        self.cb_val = ttk.Combobox(
            self, textvariable=self.v_val, width=18,
            postcommand=lambda: self.cb_val.configure(
                values=app.known_columns(self.v_alias.get()) if self._kind() == KIND_COLUMN else []))
        self.cb_val.grid(row=0, column=2, sticky="ew", **pad)
        ttk.Label(self, text="as", style="Muted.TLabel").grid(row=0, column=3, **pad)
        self.ent_out = ttk.Entry(self, textvariable=self.v_out, width=16)
        self.ent_out.grid(row=0, column=4, **pad)
        self.add_controls(self, 5)
        self.columnconfigure(2, weight=1)
        self.first_widget = self.cb_val
        self.enter_adds_row(self.ent_out)
        self.watch(self.v_alias, self.v_val, self.v_out)
        self.v_kind.trace_add("write", lambda *_: (self._sync(), self.owner.on_change()))
        self._sync()

    def _kind(self) -> str:
        return KIND_BY_LABEL.get(self.v_kind.get(), KIND_COLUMN)

    def _sync(self) -> None:
        self.cb_alias.configure(state="normal" if self._kind() == KIND_COLUMN else "disabled")

    def set_defaults(self) -> None:
        self.v_alias.set(self.app.default_alias())

    def get_item(self) -> Col:
        kind = self._kind()
        return Col(kind, self.v_alias.get().strip() if kind == KIND_COLUMN else "",
                   self.v_val.get().strip(), self.v_out.get().strip())

    def set_item(self, c: Col) -> None:
        self.v_kind.set(KIND_LABELS.get(c.kind, KIND_LABELS[KIND_COLUMN]))
        self.v_alias.set(c.table_alias)
        self.v_val.set(c.value)
        self.v_out.set(c.out_alias)
        self._sync()


class LinkRow(RowBase):
    """Join link:  ON  table . column  =  table . column"""

    def __init__(self, owner, app: "QueryBuilderApp", block: "JoinBlock"):
        super().__init__(owner)
        self.app = app
        self.block = block
        self.v_la = tk.StringVar()
        self.v_lc = tk.StringVar()
        self.v_op = tk.StringVar(value="=")
        self.v_ra = tk.StringVar()
        self.v_rc = tk.StringVar()
        pad = {"padx": (0, 4)}
        self.lbl = ttk.Label(self, text="ON", width=4, style="Muted.TLabel")
        self.lbl.grid(row=0, column=0, sticky="w", **pad)
        alias_box(self, self.v_la, app).grid(row=0, column=1, **pad)
        self.cb_lc = column_box(self, self.v_lc, self.v_la, app, 14)
        self.cb_lc.grid(row=0, column=2, sticky="ew", **pad)
        ttk.Combobox(self, textvariable=self.v_op, values=LINK_OPERATORS, state="readonly",
                     width=3).grid(row=0, column=3, **pad)
        alias_box(self, self.v_ra, app).grid(row=0, column=4, **pad)
        self.cb_rc = column_box(self, self.v_rc, self.v_ra, app, 14)
        self.cb_rc.grid(row=0, column=5, sticky="ew", **pad)
        self.add_controls(self, 6)
        self.columnconfigure(2, weight=1)
        self.columnconfigure(5, weight=1)
        self.first_widget = self.cb_lc
        self.enter_adds_row(self.cb_rc)
        self.watch(self.v_la, self.v_lc, self.v_op, self.v_ra, self.v_rc)

    def set_index(self, index: int) -> None:
        self.lbl.configure(text="ON" if index == 0 else "AND")

    def set_defaults(self) -> None:
        self.v_la.set(self.app.default_alias())
        self.v_ra.set(self.block.v_alias.get().strip())

    def get_item(self) -> Link:
        return Link(self.v_la.get().strip(), self.v_lc.get().strip(), self.v_op.get() or "=",
                    self.v_ra.get().strip(), self.v_rc.get().strip())

    def set_item(self, l: Link) -> None:
        self.v_la.set(l.left_alias)
        self.v_lc.set(l.left_col)
        self.v_op.set(l.op)
        self.v_ra.set(l.right_alias)
        self.v_rc.set(l.right_col)


class JoinBlock(RowBase):
    """One joined table: type, table, alias and its link rows."""

    def __init__(self, owner, app: "QueryBuilderApp"):
        super().__init__(owner)
        self.app = app
        self.columnconfigure(0, weight=1)
        self.v_type = tk.StringVar(value=JOIN_LABEL["left join"])
        self.v_table = tk.StringVar()
        self.v_alias = tk.StringVar()
        self._last_alias = ""

        self.box = ttk.LabelFrame(self, text="Join", padding=8)
        self.box.grid(row=0, column=0, sticky="ew")
        self.box.columnconfigure(0, weight=1)
        head = ttk.Frame(self.box)
        head.grid(row=0, column=0, sticky="ew")
        head.columnconfigure(1, weight=1)
        ttk.Combobox(head, textvariable=self.v_type, values=[l for l, _s in JOIN_LABELS],
                     state="readonly", width=16).grid(row=0, column=0, padx=(0, 6))
        self.ent_table = ttk.Entry(head, textvariable=self.v_table)
        self.ent_table.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ttk.Label(head, text="alias", style="Muted.TLabel").grid(row=0, column=2, padx=(0, 4))
        ttk.Entry(head, textvariable=self.v_alias, width=8).grid(row=0, column=3)
        self.add_controls(head, 4)
        ttk.Label(self.box, style="Muted.TLabel",
                  text="Table name only - the schema is added for you.  Then say how it links to the tables above."
                  ).grid(row=1, column=0, sticky="w", pady=(4, 2))

        self.links = RowList(self.box, lambda lst: LinkRow(lst, app, self), owner.on_change,
                             "+ Add link", see=owner.see)
        self.links.grid(row=2, column=0, sticky="ew")
        self.note = ttk.Label(self.box, style="Muted.TLabel",
                              text="A cross join combines every row with every row - no links needed.")
        self.note.grid(row=3, column=0, sticky="w")
        self.first_widget = self.ent_table

        self.v_table.trace_add("write", lambda *_: self.owner.on_change())
        self.v_alias.trace_add("write", lambda *_: self._alias_changed())
        self.v_type.trace_add("write", lambda *_: (self._sync(), self.owner.on_change()))
        self._sync()

    def _alias_changed(self) -> None:
        new, old = self.v_alias.get().strip(), self._last_alias
        if new != old:
            for row in self.links.rows:           # links that point at this join follow its alias
                if row.v_ra.get().strip() == old:
                    row.v_ra.set(new)
            self._last_alias = new
        self.owner.on_change()

    def _sync(self) -> None:
        cross = JOIN_SQL.get(self.v_type.get()) == CROSS_JOIN
        if cross:
            self.links.grid_remove()
            self.note.grid()
        else:
            self.note.grid_remove()
            self.links.grid()

    def set_index(self, index: int) -> None:
        self.box.configure(text=f"Join {index + 1}")

    def set_defaults(self) -> None:
        self.v_alias.set(self.app.next_alias())
        self.links.add(notify=False)

    def get_item(self) -> Join:
        return Join(JOIN_SQL.get(self.v_type.get(), "left join"), self.v_table.get().strip(),
                    self.v_alias.get().strip(), self.links.items())

    def set_item(self, j: Join) -> None:
        self.v_type.set(JOIN_LABEL.get(j.jtype, JOIN_LABEL["left join"]))
        self.v_table.set(j.table)
        self.v_alias.set(j.alias)
        self.links.set_items(j.links)
        self._sync()


# =============================================================================
# SECTION 3 - MAIN APPLICATION WINDOW
# =============================================================================

_KEYWORD_RE = re.compile(
    r"\b(?:insert into|select|distinct|from|inner join|left join|right join|full outer join|"
    r"cross join|on|where|and|or|not in|in|not like|like|not between|between|is not null|"
    r"is null|as)\b", re.IGNORECASE)
_STRING_RE = re.compile(r"'(?:[^']|'')*'")
_VARIABLE_TOKEN_RE = re.compile(r"\$\{[^{}\s]+\}")

HELP_TEXT = """\
HOW TO USE  -  no SQL knowledge needed

Work down the page. The SQL on the right updates by itself.
When the Checks box says "No problems found", press  Copy SQL  (or Ctrl+Enter).

1  TARGET AND SOURCE
   Schema / prefix  : added in front of every table name, e.g. ${env.schema}
   INSERT INTO      : the table the result goes into (untick for a plain SELECT)
   FROM table       : the main table, with a short name (alias) such as t1

2  JOINS  (optional)
   Press "+ Add join".  Choose the type, type the table name and a short alias.
   Then say how it links to the tables above:   t1 . id  =  t2 . id
   Use "+ Add link" when more than one column is needed.
   Left join  = keep every row of the table above, add matching rows
   Inner join = keep only rows that match on both sides

3  COLUMNS
   Press "+ Add column", choose the table and type the column name.
   "Quick add" adds many columns of one table at once:  a, b, c, d
   Output name (after "as") is optional: t1.a becomes t1_a automatically.
   Other kinds: Text constant ('abc' as abc), Number constant, SQL expression.

4  FILTERS  (optional)
   Press "+ Add filter".  Choose the table, type the column, CHOOSE how to compare,
   then type the value, for example   t1 . amt   >= at least   500
   IN / NOT IN   : type the values separated by commas:  X, Y, Z
   BETWEEN       : type the two values in the two boxes
   LIKE          : use % as a wildcard, e.g. AB%
   AND / OR      : choose how the row connects to the one above
   ( and )       : only needed to group conditions, e.g. a AND (b OR c)

VALUE TYPE (filters)
   Auto          : numbers stay numbers, text is quoted for you, t2.id is recognised as a column
   Text          : always put in quotes (use for codes such as 007)
   Number        : must be a number
   Column / SQL  : inserted exactly as typed (functions, ${variables}, sub-queries)

ROWS
   Enter in the last box adds a new row.  Use the arrows to reorder and the cross to delete.
   While you type in a Column box the dropdown lists columns you already used for that table.

SAVING AND SHARING
   Everything is stored in one Excel file, SQLQueryBuilder_saved.xlsx, next to the program.
   Your work is saved automatically and restored at the next start.
   Toolbar: type a name and press Save; pick a name and press Open or Delete.
   To share: send the program together with the Excel file.
   Close the Excel file before using the program, otherwise saving fails.
"""


class QueryBuilderApp:
    def __init__(self, root: tk.Tk, store: Optional[QueryStore] = None):
        self.root = root
        self.store = store or QueryStore()
        self.last_result: Optional[BuildResult] = None
        self._after_id = None
        self._save_after_id = None
        self._autosave_failed = False
        self._last_save_error = ""
        self._saved_snapshot = config_to_dict(QueryConfig())     # what the file holds as last session
        self._clean_snapshot = config_to_dict(QueryConfig())     # last saved / opened named query
        self.current_name = ""

        root.title(f"{APP_NAME} {APP_VERSION}")
        width = min(1320, max(1120, root.winfo_screenwidth() - 80))
        height = min(820, max(640, root.winfo_screenheight() - 110))
        root.geometry(f"{width}x{height}+{max(0, (root.winfo_screenwidth() - width) // 2)}"
                      f"+{max(0, (root.winfo_screenheight() - height) // 3)}")
        root.minsize(1060, 620)
        root.configure(background=Theme.BG)
        configure_style(root)

        self.v_schema = tk.StringVar(value=DEFAULT_SCHEMA)
        self.v_insert = tk.BooleanVar(value=True)
        self.v_target = tk.StringVar()
        self.v_from = tk.StringVar()
        self.v_from_alias = tk.StringVar(value="t1")
        self.v_distinct = tk.BooleanVar(value=False)
        self.v_upper = tk.BooleanVar(value=False)
        self.v_semicolon = tk.BooleanVar(value=False)
        self.v_prefix = tk.BooleanVar(value=True)
        self.v_saved = tk.StringVar()
        self.v_summary = tk.StringVar()
        self.v_action = tk.StringVar()
        self.v_qa_alias = tk.StringVar()
        self.v_qa_cols = tk.StringVar()
        for var in (self.v_schema, self.v_insert, self.v_target, self.v_from, self.v_from_alias,
                    self.v_distinct, self.v_upper, self.v_semicolon, self.v_prefix):
            var.trace_add("write", lambda *_: self.schedule_refresh())

        self._build_menu()
        self._build_header()
        self._build_status_bar()
        self._build_toolbar()
        self._build_body()
        self._bind_mouse_wheel()
        self.root.bind_all("<Control-Return>", lambda _e: self.copy_sql())
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        restored, problem = None, ""
        try:
            restored = self.store.load_session()
        except (OSError, StoreError) as exc:
            problem = describe_store_error(exc)
        if restored is not None:
            self.apply_config(restored)
            self._saved_snapshot = config_to_dict(restored)
            self.note("Your previous session was restored.")
        else:
            self.refresh_preview()
            if problem:
                self.note("The Excel data file could not be read.")
                messagebox.showwarning(
                    "Excel data file",
                    "The saved data could not be read, so nothing was restored and "
                    f"nothing will be saved until this is fixed:\n\n{problem}", parent=root)
        self.refresh_names()

    # ------------------------------------------------------------------ layout
    def _build_menu(self) -> None:
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Copy SQL", accelerator="Ctrl+Enter", command=self.copy_sql)
        file_menu.add_command(label="Save SQL as...", command=self.save_sql)
        file_menu.add_separator()
        file_menu.add_command(label="Where is the Excel data file?", command=self.show_data_file)
        file_menu.add_separator()
        file_menu.add_command(label="New / reset all", command=self.reset_all)
        file_menu.add_command(label="Exit", command=self.on_close)
        menubar.add_cascade(label="File", menu=file_menu)
        help_menu = tk.Menu(menubar, tearoff=False)
        help_menu.add_command(label="How to use", command=self.show_help)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menubar)

    def _build_header(self) -> None:
        header = ttk.Frame(self.root, style="Header.TFrame", padding=(18, 10))
        header.pack(fill="x")
        ttk.Label(header, text=APP_NAME, style="Header.TLabel").pack(anchor="w")
        ttk.Label(header, style="SubHeader.TLabel",
                  text="Pick tables, columns and conditions - the SQL is written for you.").pack(anchor="w")

    def _build_status_bar(self) -> None:
        bar = ttk.Frame(self.root, style="Status.TFrame", padding=(14, 5))
        bar.pack(side="bottom", fill="x")
        ttk.Label(bar, textvariable=self.v_summary, style="Status.TLabel").pack(side="left")
        ttk.Label(bar, textvariable=self.v_action, style="Status.TLabel").pack(side="right")

    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self.root, padding=(12, 8, 12, 0))
        bar.pack(fill="x")
        ttk.Label(bar, text="Saved queries:").pack(side="left")
        self.cb_saved = ttk.Combobox(bar, textvariable=self.v_saved, width=30, postcommand=self.refresh_names)
        self.cb_saved.pack(side="left", padx=(6, 6))
        ttk.Button(bar, text="Save", command=self.save_named).pack(side="left")
        ttk.Button(bar, text="Open", command=self.open_named).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Delete", command=self.delete_named).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="New", command=self.reset_all).pack(side="right")
        ttk.Button(bar, text="Load example", command=self.load_example).pack(side="right", padx=(0, 6))

    def _section(self, title: str, intro: str) -> ttk.LabelFrame:
        box = ttk.LabelFrame(self.scroll.inner, text=title, padding=(10, 8))
        box.grid(row=self._section_row, column=0, sticky="ew", pady=(0, 10), padx=(0, 6))
        box.columnconfigure(0, weight=1)
        self._section_row += 1
        if intro:
            ttk.Label(box, text=intro, style="Muted.TLabel", wraplength=640, justify="left").grid(
                row=0, column=0, sticky="w", pady=(0, 6))
        return box

    def _build_body(self) -> None:
        paned = ttk.PanedWindow(self.root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=12, pady=(8, 8))
        left, right = ttk.Frame(paned), ttk.Frame(paned)
        paned.add(left, weight=4)
        paned.add(right, weight=3)

        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)
        self.scroll = ScrollPage(left)
        self.scroll.grid(row=0, column=0, sticky="nsew")
        self.scroll.inner.columnconfigure(0, weight=1)
        self._section_row = 0

        # ---- 1 target and source
        s1 = self._section("1   Where does the data go and where does it come from?", "")
        s1.columnconfigure(1, weight=1)
        ttk.Label(s1, text="Schema / prefix").grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Entry(s1, textvariable=self.v_schema).grid(row=0, column=1, sticky="ew")
        ttk.Checkbutton(s1, text="INSERT INTO", variable=self.v_insert,
                        command=self._sync_target_state).grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.ent_target = ttk.Entry(s1, textvariable=self.v_target)
        self.ent_target.grid(row=1, column=1, sticky="ew", pady=(8, 0))
        ttk.Label(s1, text="FROM table").grid(row=2, column=0, sticky="w", pady=(8, 0), padx=(0, 10))
        src = ttk.Frame(s1)
        src.grid(row=2, column=1, sticky="ew", pady=(8, 0))
        src.columnconfigure(0, weight=1)
        ttk.Entry(src, textvariable=self.v_from).grid(row=0, column=0, sticky="ew")
        ttk.Label(src, text="alias", style="Muted.TLabel").grid(row=0, column=1, padx=(8, 4))
        ttk.Entry(src, textvariable=self.v_from_alias, width=8).grid(row=0, column=2)
        ttk.Label(s1, style="Muted.TLabel", justify="left",
                  text="Type table names only (e.g. table1) - the schema is added automatically.\n"
                       "A name that already contains a dot (other_db.table) is used exactly as typed."
                  ).grid(row=3, column=1, sticky="w", pady=(4, 0))
        opts = ttk.Frame(s1)
        opts.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(opts, text="Name columns t1_a (otherwise a)", variable=self.v_prefix).pack(side="left")
        ttk.Checkbutton(opts, text="SELECT DISTINCT", variable=self.v_distinct).pack(side="left", padx=(14, 0))
        ttk.Checkbutton(opts, text="UPPERCASE keywords", variable=self.v_upper).pack(side="left", padx=(14, 0))
        ttk.Checkbutton(opts, text="End with ;", variable=self.v_semicolon).pack(side="left", padx=(14, 0))

        # ---- 2 joins
        s2 = self._section("2   Join other tables  (optional)",
                           "Left join keeps every row of the tables above and adds the matching rows. "
                           "Inner join keeps only rows that match.")
        self.join_list = RowList(s2, lambda lst: JoinBlock(lst, self), self.schedule_refresh,
                                 "+ Add join", see=self.scroll.see)
        self.join_list.grid(row=1, column=0, sticky="ew")

        # ---- 3 columns
        s3 = self._section("3   Which columns do you want?",
                           "Choose the table and type the column name. The output name (after 'as') is optional.")
        self.col_list = RowList(s3, lambda lst: ColRow(lst, self), self.schedule_refresh,
                                "+ Add column", see=self.scroll.see)
        self.col_list.grid(row=1, column=0, sticky="ew")
        quick = ttk.Frame(s3)
        quick.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        quick.columnconfigure(2, weight=1)
        ttk.Label(quick, text="Quick add", style="Muted.TLabel").grid(row=0, column=0, padx=(0, 6))
        alias_box(quick, self.v_qa_alias, self).grid(row=0, column=1, padx=(0, 6))
        qa_entry = ttk.Entry(quick, textvariable=self.v_qa_cols)
        qa_entry.grid(row=0, column=2, sticky="ew", padx=(0, 6))
        qa_entry.bind("<Return>", lambda _e: self.quick_add())
        ttk.Button(quick, text="Add columns", command=self.quick_add).grid(row=0, column=3)
        ttk.Label(quick, text="Several columns of one table, separated by commas:  a, b, c, d",
                  style="Muted.TLabel").grid(row=1, column=2, sticky="w", pady=(2, 0))

        # ---- 4 filters
        s4 = self._section("4   Filter the rows  (optional)",
                           "Choose the table, type the column, choose how to compare, then type the value. "
                           "Example:  t1 . amt   >= at least   500")
        self.filter_list = RowList(s4, lambda lst: CondRow(lst, self), self.schedule_refresh,
                                   "+ Add filter", see=self.scroll.see)
        self.filter_list.grid(row=1, column=0, sticky="ew")

        self._build_preview(right)
        self._sync_target_state()

    def _build_preview(self, right) -> None:
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=3)
        right.rowconfigure(1, weight=1)

        sql_box = ttk.LabelFrame(right, text="Generated SQL", padding=8)
        sql_box.grid(row=0, column=0, sticky="nsew")
        sql_box.columnconfigure(0, weight=1)
        sql_box.rowconfigure(0, weight=1)
        self.txt_sql = tk.Text(sql_box, wrap="none", height=12, width=36, font="TkFixedFont",
                               background=Theme.EDITOR_BG, foreground=Theme.EDITOR_FG,
                               insertbackground=Theme.EDITOR_FG, relief="flat", padx=10, pady=8,
                               state="disabled", borderwidth=0, highlightthickness=0)
        ysb = ttk.Scrollbar(sql_box, orient="vertical", command=self.txt_sql.yview)
        xsb = ttk.Scrollbar(sql_box, orient="horizontal", command=self.txt_sql.xview)
        self.txt_sql.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.txt_sql.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        self.txt_sql.tag_configure("placeholder", foreground=Theme.EDITOR_MUTED)
        self.txt_sql.tag_configure("kw", foreground=Theme.EDITOR_KEYWORD)
        self.txt_sql.tag_configure("str", foreground=Theme.EDITOR_STRING)
        self.txt_sql.tag_configure("var", foreground=Theme.EDITOR_VARIABLE)

        check_box = ttk.LabelFrame(right, text="Checks", padding=8)
        check_box.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        check_box.columnconfigure(0, weight=1)
        check_box.rowconfigure(0, weight=1)
        self.txt_issues = tk.Text(check_box, wrap="word", height=6, width=36, background=Theme.CARD,
                                  foreground=Theme.TEXT, relief="flat", padx=6, pady=4,
                                  state="disabled", borderwidth=0, highlightthickness=0)
        isb = ttk.Scrollbar(check_box, orient="vertical", command=self.txt_issues.yview)
        self.txt_issues.configure(yscrollcommand=isb.set)
        self.txt_issues.grid(row=0, column=0, sticky="nsew")
        isb.grid(row=0, column=1, sticky="ns")
        self.txt_issues.tag_configure("error", foreground=Theme.ERROR)
        self.txt_issues.tag_configure("warning", foreground=Theme.WARNING)
        self.txt_issues.tag_configure("ok", foreground=Theme.OK)

        buttons = ttk.Frame(right)
        buttons.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.btn_copy = ttk.Button(buttons, text="Copy SQL  (Ctrl+Enter)", style="Accent.TButton",
                                   command=self.copy_sql)
        self.btn_copy.pack(side="left")
        self.btn_save = ttk.Button(buttons, text="Save SQL...", command=self.save_sql)
        self.btn_save.pack(side="left", padx=(8, 0))

    def _bind_mouse_wheel(self) -> None:
        """Mouse wheel scrolls the page - also when the pointer is over a dropdown (which would
        otherwise change its value by accident)."""
        def on_wheel(event):
            target = self.root.winfo_containing(event.x_root, event.y_root)
            if target is None or not str(target).startswith(str(self.scroll.canvas)):
                return None
            if getattr(event, "num", 0) == 4:
                direction = -1
            elif getattr(event, "num", 0) == 5:
                direction = 1
            else:
                direction = -1 if event.delta > 0 else 1
            self.scroll.scroll(direction)
            return "break"

        def on_combo_wheel(event):
            on_wheel(event)
            return "break"

        for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.root.bind_all(sequence, on_wheel)
            self.root.bind_class("TCombobox", sequence, on_combo_wheel)

    # ------------------------------------------------------------ model <-> GUI
    def get_aliases(self) -> List[str]:
        """Table aliases defined so far: FROM alias first, then one per join."""
        names = [self.v_from_alias.get().strip()]
        join_list = getattr(self, "join_list", None)
        if join_list is not None:
            names += [r.v_alias.get().strip() for r in join_list.rows]
        unique: List[str] = []
        for name in names:
            if name and name not in unique:
                unique.append(name)
        return unique

    def default_alias(self) -> str:
        aliases = self.get_aliases()
        return aliases[0] if aliases else ""

    def next_alias(self) -> str:
        used = {a.lower() for a in self.get_aliases()}
        n = 1
        while f"t{n}" in used:
            n += 1
        return f"t{n}"

    def known_columns(self, alias: str) -> List[str]:
        """Columns already used for this table alias anywhere in the query (for the dropdowns)."""
        cfg = self.collect_config()
        wanted = alias.strip().lower()
        found: List[str] = []

        def add(a: str, column: str) -> None:
            column = column.strip()
            if column and column != "*" and a.strip().lower() == wanted and column not in found:
                found.append(column)

        for item in cfg.select_items:
            if item.kind == KIND_COLUMN:
                add(item.table_alias, item.value)
        for cond in cfg.filters:
            add(cond.alias, cond.column)
        for join in cfg.joins:
            for link in join.links:
                add(link.left_alias, link.left_col)
                add(link.right_alias, link.right_col)
        return found

    def collect_config(self) -> QueryConfig:
        return QueryConfig(
            schema=self.v_schema.get(), use_insert=self.v_insert.get(),
            target_table=self.v_target.get(), from_table=self.v_from.get(),
            from_alias=self.v_from_alias.get(), joins=self.join_list.items(),
            select_items=self.col_list.items(), filters=self.filter_list.items(),
            distinct=self.v_distinct.get(), uppercase=self.v_upper.get(),
            semicolon=self.v_semicolon.get(), prefix_names=self.v_prefix.get())

    def apply_config(self, cfg: QueryConfig) -> None:
        self.v_schema.set(cfg.schema)
        self.v_insert.set(cfg.use_insert)
        self.v_target.set(cfg.target_table)
        self.v_from.set(cfg.from_table)
        self.v_from_alias.set(cfg.from_alias)
        self.v_distinct.set(cfg.distinct)
        self.v_upper.set(cfg.uppercase)
        self.v_semicolon.set(cfg.semicolon)
        self.v_prefix.set(cfg.prefix_names)
        self.join_list.set_items(cfg.joins)
        self.col_list.set_items(cfg.select_items)
        self.filter_list.set_items(cfg.filters)
        self.v_qa_alias.set(self.default_alias())
        self._sync_target_state()
        self.refresh_preview()

    def _sync_target_state(self) -> None:
        self.ent_target.configure(state="normal" if self.v_insert.get() else "disabled")

    def quick_add(self) -> None:
        alias = self.v_qa_alias.get().strip() or self.default_alias()
        names = [n for n in re.split(r"[,\s]+", self.v_qa_cols.get()) if n]
        parent = self.root
        if not alias:
            messagebox.showwarning("Please check your input", "Choose the table first.", parent=parent)
            return
        if not names:
            messagebox.showwarning("Please check your input",
                                   "Type at least one column name, separated by commas.", parent=parent)
            return
        for name in names:
            try:
                select_item_parts(Col(KIND_COLUMN, alias, name, ""), self.v_prefix.get())
            except ValueError as exc:
                messagebox.showwarning("Please check your input",
                                       sentence(f"{name}: {exc}") + "\nNothing was added.", parent=parent)
                return
        for name in names:
            self.col_list.add(Col(KIND_COLUMN, alias, name, ""), notify=False)
        self.v_qa_cols.set("")
        self.schedule_refresh()
        self.scroll.see(self.col_list.rows[-1])

    # ---------------------------------------------------------------- preview
    def schedule_refresh(self) -> None:
        if self._after_id is not None:
            self.root.after_cancel(self._after_id)
        self._after_id = self.root.after(150, self.refresh_preview)

    def refresh_preview(self) -> None:
        self._after_id = None
        result = build_query(self.collect_config())
        self.last_result = result
        self._schedule_autosave()
        self._show_sql(result)
        self._show_issues(result)
        ok = result.sql is not None
        state = "normal" if ok else "disabled"
        self.btn_copy.configure(state=state)
        self.btn_save.configure(state=state)
        if ok:
            w = len(result.warnings)
            self.v_summary.set("\u2714 Query is ready" + (f"  ({w} warning{'s' if w != 1 else ''})" if w else ""))
        else:
            n = len(result.errors)
            self.v_summary.set(f"\u2716 {n} thing{'s' if n != 1 else ''} to fix before the query can be copied")

    def _show_sql(self, result: BuildResult) -> None:
        txt = self.txt_sql
        txt.configure(state="normal")
        txt.delete("1.0", "end")
        for tag in ("placeholder", "kw", "str", "var"):
            txt.tag_remove(tag, "1.0", "end")
        if result.sql is None:
            txt.insert("1.0", "-- The query is not complete yet.\n-- Fix the items listed under 'Checks' below.",
                       "placeholder")
        else:
            txt.insert("1.0", result.sql)
            for tag, rx in (("kw", _KEYWORD_RE), ("str", _STRING_RE), ("var", _VARIABLE_TOKEN_RE)):
                for m in rx.finditer(result.sql):
                    txt.tag_add(tag, f"1.0+{m.start()}c", f"1.0+{m.end()}c")
        txt.configure(state="disabled")

    def _show_issues(self, result: BuildResult) -> None:
        txt = self.txt_issues
        txt.configure(state="normal")
        txt.delete("1.0", "end")
        if not result.issues:
            txt.insert("end", "\u2714 No problems found. The query is ready to copy.", "ok")
        else:
            for issue in result.errors:
                txt.insert("end", f"\u2716  {issue.message}\n", "error")
            for issue in result.warnings:
                txt.insert("end", f"\u26A0  {issue.message}\n", "warning")
        txt.configure(state="disabled")

    # ---------------------------------------------------------------- autosave
    def _schedule_autosave(self) -> None:
        if self._save_after_id is not None:
            self.root.after_cancel(self._save_after_id)
        self._save_after_id = self.root.after(AUTOSAVE_DELAY_MS, self._autosave_now)

    def _autosave_now(self) -> bool:
        self._save_after_id = None
        cfg = self.collect_config()
        snapshot = config_to_dict(cfg)
        if snapshot == self._saved_snapshot:
            return True
        try:
            self.store.save_session(cfg)
        except (OSError, StoreError) as exc:
            self._last_save_error = describe_store_error(exc)
            if not self._autosave_failed:
                self._autosave_failed = True
                self.note("Could not save to the Excel file (is it open in Excel?).")
            return False
        self._saved_snapshot = snapshot
        self._autosave_failed = False
        self._last_save_error = ""
        return True

    def on_close(self) -> None:
        if self._save_after_id is not None:
            self.root.after_cancel(self._save_after_id)
            self._save_after_id = None
        if not self._autosave_now():
            messagebox.showwarning(
                "Excel data file",
                f"Your latest changes could not be saved to the Excel file:\n\n{self._last_save_error}",
                parent=self.root)
        self.root.destroy()

    # ---------------------------------------------------------- saved queries
    def refresh_names(self) -> None:
        try:
            names = self.store.list_queries()
        except (OSError, StoreError) as exc:
            names = []
            self.note("Could not read the Excel file: " + describe_store_error(exc).splitlines()[0])
        self.cb_saved.configure(values=names)

    def _has_unsaved_work(self) -> bool:
        snap = config_to_dict(self.collect_config())
        return snap != self._clean_snapshot and snap != config_to_dict(QueryConfig())

    def save_named(self) -> None:
        parent = self.root
        try:
            name = QueryStore.clean_name(self.v_saved.get())
            if name != self.current_name and self.store.has_query(name) and not messagebox.askyesno(
                    "Replace saved query", f"A query named '{name}' already exists. Replace it?", parent=parent):
                return
            cfg = self.collect_config()
            self.store.save_query(name, cfg)
        except ValueError as exc:
            messagebox.showwarning("Please check your input", sentence(str(exc)), parent=parent)
            return
        except (OSError, StoreError) as exc:
            messagebox.showerror("Could not save",
                                 f"The query could not be saved:\n{describe_store_error(exc)}", parent=parent)
            return
        self.v_saved.set(name)
        self.current_name = name
        self._clean_snapshot = config_to_dict(cfg)
        self.refresh_names()
        self.note(f"Saved '{name}'.")

    def open_named(self) -> None:
        name = " ".join(self.v_saved.get().split())
        if not name:
            messagebox.showinfo("Open", "Pick or type the name of a saved query first.", parent=self.root)
            return
        if self._has_unsaved_work() and not messagebox.askyesno(
                "Open saved query",
                f"The query on screen has not been saved under a name.\nOpen '{name}' anyway and replace it?",
                parent=self.root):
            return
        try:
            cfg = self.store.load_query(name)
        except (OSError, StoreError, ValueError) as exc:
            messagebox.showerror("Could not open",
                                 f"The query could not be loaded:\n{describe_store_error(exc)}", parent=self.root)
            return
        self.apply_config(cfg)
        self.current_name = name
        self._clean_snapshot = config_to_dict(self.collect_config())
        self.note(f"Opened '{name}'.")

    def delete_named(self) -> None:
        name = " ".join(self.v_saved.get().split())
        if not name:
            messagebox.showinfo("Delete", "Pick the saved query to delete first.", parent=self.root)
            return
        if not messagebox.askyesno("Delete saved query", f"Delete '{name}'? This cannot be undone.",
                                   parent=self.root):
            return
        try:
            self.store.delete_query(name)
        except (OSError, StoreError) as exc:
            messagebox.showerror("Could not delete",
                                 f"The query could not be deleted:\n{describe_store_error(exc)}", parent=self.root)
            return
        if name == self.current_name:
            self.current_name = ""
        self.v_saved.set("")
        self.refresh_names()
        self.note(f"Deleted '{name}'.")

    # ---------------------------------------------------------------- actions
    def note(self, message: str) -> None:
        self.v_action.set(message)
        self.root.after(5000, lambda: self.v_action.set("") if self.v_action.get() == message else None)

    def _current_sql(self) -> Optional[str]:
        return self.last_result.sql if self.last_result is not None else None

    def copy_sql(self) -> None:
        sql = self._current_sql()
        if sql is None:
            self.note("Fix the problems listed under 'Checks' first.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(sql)
        self.root.update_idletasks()
        self.note("SQL copied to the clipboard.")

    def save_sql(self) -> None:
        sql = self._current_sql()
        if sql is None:
            self.note("Fix the problems listed under 'Checks' first.")
            return
        path = filedialog.asksaveasfilename(
            parent=self.root, title="Save SQL", defaultextension=".sql",
            filetypes=[("SQL files", "*.sql"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8", newline="\n") as handle:
                handle.write(sql + "\n")
        except OSError as exc:
            messagebox.showerror("Could not save", f"The file could not be written:\n{exc}", parent=self.root)
            return
        self.note(f"Saved {path}")

    def load_example(self) -> None:
        if self._has_unsaved_work() and not messagebox.askyesno(
                "Load example", "Replace the query on screen with the example?", parent=self.root):
            return
        self.apply_config(EXAMPLE)
        self.current_name = ""

    def reset_all(self) -> None:
        if self._has_unsaved_work() and not messagebox.askyesno(
                "New query", "Clear everything and start a new query?", parent=self.root):
            return
        self.apply_config(QueryConfig())
        self.current_name = ""
        self.v_saved.set("")
        self._clean_snapshot = config_to_dict(QueryConfig())
        self.note("Started a new query.")

    def show_data_file(self) -> None:
        exists = os.path.exists(self.store.path)
        messagebox.showinfo(
            "Excel data file",
            "Saved queries and your last session are stored in this Excel file:\n\n"
            f"{self.store.path}\n\n"
            + ("" if exists else "(It does not exist yet - it is created the first time something is saved.)\n\n")
            + "Send it together with the program to share your queries.", parent=self.root)

    def show_help(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("How to use")
        win.geometry("760x680")
        win.transient(self.root)
        text = tk.Text(win, wrap="word", padx=14, pady=12, font="TkDefaultFont", relief="flat",
                       background=Theme.CARD, foreground=Theme.TEXT, borderwidth=0, highlightthickness=0)
        scroll = ttk.Scrollbar(win, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        text.pack(side="left", fill="both", expand=True)
        text.insert("1.0", HELP_TEXT + f"\nData file:\n   {self.store.path}\n")
        text.configure(state="disabled")

    def show_about(self) -> None:
        messagebox.showinfo(
            "About",
            f"{APP_NAME} {APP_VERSION}\n\nBuilds INSERT ... SELECT statements from simple choices.\n"
            "Written with Python and Tkinter (standard library only).", parent=self.root)


def main() -> None:
    enable_high_dpi()
    root = tk.Tk()
    QueryBuilderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
