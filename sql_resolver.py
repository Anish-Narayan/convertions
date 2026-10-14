"""Statement-level table/column resolution on top of SQLGlot scopes (shared by analysis AND conversion).

Every ColumnRef/TableRef keeps a reference to the live AST node so the converter can edit exactly that node.
Column status: resolved | derived (CTE/subquery output) | alias (select alias) | star | ambiguous | unresolved
"""
from dataclasses import dataclass, field
import sqlglot
from sqlglot import exp
from sqlglot.optimizer.scope import traverse_scope
from sql_common import DIALECT, table_key

SETOPS = tuple({exp.Union, getattr(exp, "SetOperation", exp.Union)})
QUERY_TYPES = (exp.Select, exp.Subquery) + SETOPS
OWNER_TYPES = (exp.Select,) + SETOPS
ROLES = {exp.Create: "create", exp.Insert: "insert", exp.Drop: "drop",
         exp.Update: "update", exp.Delete: "delete", exp.Merge: "merge"}

# legacy operation labels (kept from the original script)
OPERATION_TYPES = [(exp.Select, "SELECT"), (exp.Insert, "INSERT"), (exp.Update, "UPDATE"),
                   (exp.Delete, "DELETE"), (exp.Create, "CREATE"), (exp.Drop, "DROP"),
                   (exp.Where, "WHERE"), (exp.Group, "GROUP_BY"), (exp.Order, "ORDER_BY"),
                   (exp.Having, "HAVING"), (exp.Union, "UNION"), (exp.CTE, "CTE")]


@dataclass(eq=False)
class TableRef:
    node: object
    key: str
    alias: str
    role: str            # read | create | insert | update | delete | merge | drop
    internal: bool       # created earlier (or defined) inside this concept
    statement: int
    renamed_to: tuple = None
    pair: tuple = None


@dataclass(eq=False)
class ColumnRef:
    node: object
    name: str
    statement: int
    status: str
    table: TableRef = None
    kind: str = "column"     # column | insert_target
    qualifier: str = ""


@dataclass(eq=False)
class StatementInfo:
    index: int
    kind: str
    node: object
    tables: list = field(default_factory=list)
    columns: list = field(default_factory=list)
    joins: list = field(default_factory=list)
    ctes: list = field(default_factory=list)
    aliases: list = field(default_factory=list)
    operations: list = field(default_factory=list)
    subqueries: int = 0
    created: str = None
    created_columns: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


@dataclass
class ScriptState:
    created: dict = field(default_factory=dict)       # key -> statement index (so far)
    all_created: set = field(default_factory=set)     # every key created anywhere in the concept
    created_columns: dict = field(default_factory=dict)


def _target_table(stmt):
    if type(stmt) not in ROLES:
        return None
    this = stmt.args.get("this")
    if isinstance(this, exp.Schema):
        this = this.this
    return this if isinstance(this, exp.Table) else None


def _within(node, stop, types):
    p = node.parent
    while p is not None and p is not stop:
        if isinstance(p, types):
            return True
        p = p.parent
    return False


def rename_node(node, new):
    """Rename a Column / Identifier / ColumnDef in place."""
    ident = exp.to_identifier(new)
    if isinstance(node, exp.Identifier):
        node.replace(ident)
    else:
        node.set("this", ident)


def resolve_statement(stmt, idx, state):
    info = StatementInfo(index=idx, kind=type(stmt).__name__.upper(), node=stmt)
    if isinstance(stmt, exp.Command):
        info.warnings.append("Statement type not parsed by SQLGlot (Command); kept verbatim, not analysed.")
    info.operations = sorted({lbl for t, lbl in OPERATION_TYPES if any(True for _ in stmt.find_all(t))})
    cte_names = {c.alias_or_name.lower() for c in stmt.find_all(exp.CTE) if c.find_ancestor(*QUERY_TYPES) is None}
    info.ctes = [c.alias_or_name for c in stmt.find_all(exp.CTE)]
    info.subqueries = sum(1 for _ in stmt.find_all(exp.Subquery))
    role = ROLES.get(type(stmt))
    target = _target_table(stmt)

    def is_cte_ref(t):
        return not t.db and t.name.lower() in cte_names

    def real(t):
        return isinstance(t.this, exp.Identifier) and not is_cte_ref(t)

    # 1. tables outside any query (DDL/DML targets, UPDATE ... FROM etc.)
    top = {}
    for t in stmt.find_all(exp.Table):
        if t.find_ancestor(*QUERY_TYPES) is not None or not real(t):
            continue
        key = table_key(t)
        if t is target:
            if role == "create":
                internal = True
            elif role in ("insert", "update", "delete", "merge"):
                internal = key in state.created
            else:
                internal = key in state.all_created
            ref = TableRef(t, key, t.alias, role, internal, idx)
        else:
            ref = TableRef(t, key, t.alias, "read", key in state.created, idx)
        top[id(t)] = ref
        info.tables.append(ref)

    # 2. scopes for every outermost query (CTEs, subqueries, derived tables)
    scopes, scope_of, qrefs = [], {}, {}
    roots = [n for n in stmt.find_all(*QUERY_TYPES) if n.find_ancestor(*QUERY_TYPES) is None]
    for r in roots:
        try:
            scopes.extend(traverse_scope(r))
        except Exception as e:  # noqa
            info.errors.append(f"Scope analysis failed: {e}")
    scope_of = {id(s.expression): s for s in scopes}
    try:
        for s in scopes:
            for _, (_, source) in s.selected_sources.items():
                if isinstance(source, exp.Table) and real(source) and id(source) not in qrefs:
                    key = table_key(source)
                    ref = TableRef(source, key, source.alias, "read", key in state.created, idx)
                    qrefs[id(source)] = ref
                    info.tables.append(ref)
    except Exception as e:  # noqa
        info.errors.append(f"Source analysis failed: {e}")
    refs = {**top, **qrefs}
    flat = {r.node.alias_or_name.lower(): r for r in top.values()}

    def from_source(source):
        if isinstance(source, exp.Table):
            if not real(source):
                return "derived", None
            ref = refs.get(id(source))
            return ("resolved", ref) if ref else ("unresolved", None)
        return "derived", None

    def lookup(scope, q):
        q, s = q.lower(), scope
        while s is not None:
            for name, (_, source) in s.selected_sources.items():
                if name.lower() == q:
                    return source
            s = s.parent
        return None

    def resolve(col):
        name, qual = col.name, col.table
        owner = col.find_ancestor(*OWNER_TYPES)
        scope = scope_of.get(id(owner)) if owner is not None else None
        if owner is not None and scope is None:
            return ColumnRef(col, name, idx, "derived", qualifier=qual)
        star = col.is_star
        if scope is None:                                   # flat DML context
            if qual:
                ref = flat.get(qual.lower())
                st, ref = ("resolved", ref) if ref else ("unresolved", None)
            elif len(flat) == 1:
                st, ref = "resolved", next(iter(flat.values()))
            else:
                st, ref = "ambiguous", None
        elif qual:
            src = lookup(scope, qual)
            st, ref = from_source(src) if src is not None else ("unresolved", None)
        else:
            aliases = {a.alias.lower() for a in getattr(scope.expression, "expressions", []) if isinstance(a, exp.Alias)}
            if name.lower() in aliases and _within(col, owner, (exp.Order, exp.Group, exp.Having, exp.Qualify)):
                return ColumnRef(col, name, idx, "alias")
            srcs = [s for _, s in scope.selected_sources.values()]
            if len(srcs) == 1:
                st, ref = from_source(srcs[0])
            else:
                st, ref = ("ambiguous" if srcs else "unresolved"), None
        if star:
            st = "star" if st in ("resolved", "derived") else st
        return ColumnRef(col, name, idx, st, ref, qualifier=qual)

    for col in stmt.find_all(exp.Column):
        info.columns.append(resolve(col))
    by_node = {id(c.node): c for c in info.columns}

    # 3. INSERT INTO t (a, b) ...  - column list belongs to the target table
    if role == "insert" and target is not None and isinstance(stmt.args.get("this"), exp.Schema):
        for e in stmt.args["this"].expressions:
            info.columns.append(ColumnRef(e, e.name, idx, "resolved", top.get(id(target)), kind="insert_target"))

    # 4. joins
    for j in stmt.find_all(exp.Join):
        jt = " ".join(str(j.args[k]) for k in ("method", "side", "kind") if j.args.get(k))
        on, using = j.args.get("on"), j.args.get("using") or []
        jt = jt or ("INNER" if on is not None or using else "CROSS/COMMA")
        cols = [[(by_node[id(c)].table.key if id(c) in by_node and by_node[id(c)].table else by_node[id(c)].status), c.name]
                for c in (on.find_all(exp.Column) if on is not None else []) if id(c) in by_node]
        info.joins.append({"type": jt, "joined": getattr(j.this, "alias_or_name", ""),
                           "on": on.sql(dialect=DIALECT) if on is not None else None,
                           "using": [u.name for u in using], "columns": cols})
    info.aliases = [{"alias": a.alias, "expression": a.this.sql(dialect=DIALECT)}
                    for a in stmt.find_all(exp.Alias) if isinstance(a.parent, exp.Select) and a.arg_key == "expressions"]

    # 5. register tables created here (after reads, so CREATE t AS SELECT .. FROM t is not self-internal)
    if role == "create" and target is not None:
        key = table_key(target)
        cols = []
        this = stmt.args.get("this")
        if isinstance(this, exp.Schema):
            cols = [e.name for e in this.expressions if isinstance(e, exp.ColumnDef)]
        elif stmt.args.get("expression") is not None:
            try:
                cols = list(stmt.args["expression"].named_selects)
            except Exception:  # noqa
                pass
        info.created, info.created_columns = key, cols
        state.created[key] = idx
        state.created_columns[key] = cols
    return info


def resolve_script(statements):
    state = ScriptState()
    for s in statements:
        t = _target_table(s) if isinstance(s, exp.Create) else None
        if t is not None and str(s.args.get("kind") or "").upper() in ("TABLE", "VIEW"):
            state.all_created.add(table_key(t))
    infos = []
    for i, s in enumerate(statements, 1):
        try:
            infos.append(resolve_statement(s, i, state))
        except Exception as e:  # noqa
            info = StatementInfo(index=i, kind=type(s).__name__.upper(), node=s)
            info.errors.append(f"Resolution failed: {e}")
            infos.append(info)
    return infos, state


# ---- legacy fingerprint (unchanged logic) -----------------------------------------------------
def fingerprint(statement):
    def anonymize(node):
        if isinstance(node, exp.Table):
            node.set("this", exp.to_identifier("TABLE_NAME"))
            node.set("db", None)
            node.set("catalog", None)
        elif isinstance(node, exp.Column):
            node.set("this", exp.to_identifier("COLUMN_NAME"))
        elif isinstance(node, exp.Literal):
            node.replace(exp.Var(this="LITERAL"))
        return node
    return statement.copy().transform(anonymize).sql(dialect=DIALECT).lower()
