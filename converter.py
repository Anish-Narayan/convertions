"""Component B engine - converts SQL from one pod to another using a mapping workbook. Never executes SQL."""
from collections import defaultdict
from dataclasses import dataclass, field

import sqlglot
from sqlglot import exp

from mapping_book import MappingBook
from sql_common import DIALECT, SchemaRegistry, lookup_entries, match_entry, norm_key, split_parts
from sql_resolver import QUERY_TYPES, rename_node, resolve_script


@dataclass
class ConversionResult:
    ok: bool = False
    sql: str = ""
    warnings: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    applied: list = field(default_factory=list)
    summary: dict = field(default_factory=lambda: dict(
        tables_mapped=0, columns_mapped=0, join_references_updated=0,
        unmapped_objects=0, ambiguous_mappings=0, aliases_added=0))

    def report(self):
        s = self.summary
        lines = ["CONVERSION " + ("SUCCEEDED" if self.ok else "FAILED"), "",
                 f"Tables mapped:            {s['tables_mapped']}",
                 f"Columns mapped:           {s['columns_mapped']}",
                 f"Join references updated:  {s['join_references_updated']}",
                 f"Unmapped objects:         {s['unmapped_objects']}",
                 f"Ambiguous mappings:       {s['ambiguous_mappings']}",
                 f"Aliases added (to keep internal column names stable): {s['aliases_added']}"]
        for title, items in (("ERRORS", self.errors), ("WARNINGS", self.warnings), ("APPLIED", self.applied)):
            if items:
                lines += ["", title + ":"] + [f"  - {i}" for i in items]
        lines += ["", "Note: mappings are name/mapping-file based; semantic equivalence is NOT validated. Review before use."]
        return "\n".join(lines)


def _fail(res, msg):
    res.ok, res.sql = False, ""
    res.errors.append(msg)
    return res


def _in_join_condition(col):
    return isinstance(col.find_ancestor(exp.Join, *QUERY_TYPES), exp.Join)


def _apply_table(ref, target_raw, schema_map, kept_schemas):
    node = ref.node
    old_cat, old_db = node.catalog, node.db
    parts = split_parts(target_raw)
    new_name = parts[-1]
    if len(parts) >= 2:
        new_db, new_cat = parts[-2], (parts[-3] if len(parts) >= 3 else None)
    else:
        new_db, new_cat = old_db or None, old_cat or None
        mapped = schema_map.get(old_db.lower()) if old_db else None
        if mapped:
            mp = split_parts(mapped)
            new_db, new_cat = mp[-1], (mp[-2] if len(mp) >= 2 else new_cat)
        elif old_db:
            kept_schemas.add(old_db)
    node.set("this", exp.to_identifier(new_name))
    node.set("db", exp.to_identifier(new_db) if new_db else None)
    node.set("catalog", exp.to_identifier(new_cat) if new_cat else None)
    ref.renamed_to = (new_cat, new_db, new_name)


def _requalify(col, ref):
    """Column written as claims.col (no alias) must follow the table rename."""
    if not col.table or ref.alias:
        return False
    cat, db, name = ref.renamed_to
    col.set("table", exp.to_identifier(name))
    if col.args.get("db") is not None:
        col.set("db", exp.to_identifier(db) if db else None)
    if col.args.get("catalog") is not None:
        col.set("catalog", exp.to_identifier(cat) if cat else None)
    return True


def convert_sql(sql, book: MappingBook, source_label, target_label, preserve_internal_names=True, force=False):
    res = ConversionResult()
    S = res.summary
    if not sql.strip():
        return _fail(res, "Source SQL is empty.")
    try:
        d = book.directed(source_label, target_label)
    except ValueError as e:
        return _fail(res, str(e))
    if not d.table_rows:
        return _fail(res, "The workbook contains no table mappings.")

    # --- 1-3: schema placeholders -> temp ids (same logic as the analysis engine) ---
    registry = SchemaRegistry()
    try:
        tmp = registry.replace(sql)
        entries, tgt_keys, colmap, schema_map = defaultdict(list), set(), defaultdict(lambda: defaultdict(list)), {}
        for s_raw, t_raw in d.table_rows:
            t = registry.replace(t_raw)
            entries[norm_key(registry.replace(s_raw))].append(t)
            tgt_keys.add(norm_key(t))
        for r in d.column_rows:
            if r.get("source_column") and r.get("target_column"):
                pair = (norm_key(registry.replace(r["source_table"])), norm_key(registry.replace(r["target_table"])))
                colmap[pair][r["source_column"].lower()].append((r["target_column"], (r.get("status") or "confirmed").lower()))
        for a, b in d.schema_rows:
            schema_map[norm_key(registry.replace(a))] = registry.replace(b)
    except ValueError as e:
        return _fail(res, str(e))
    if not d.column_rows:
        res.warnings.append("Workbook has no column mappings: only table names will be converted.")

    # --- 4: parse (fail loudly) ---
    try:
        stmts = [s for s in sqlglot.parse(tmp, read=DIALECT) if s is not None]
    except Exception as e:  # noqa
        return _fail(res, f"Parse error ({DIALECT}): {e}")
    if not stmts:
        return _fail(res, "No SQL statements found.")
    infos, _ = resolve_script(stmts)
    for i in infos:
        res.errors += [f"Statement {i.index}: {e}" for e in i.errors]
        res.warnings += [f"Statement {i.index}: {w}" for w in i.warnings]
    if res.errors:
        return _fail(res, "Analysis failed; no output generated.")

    # --- safeguard: already converted? ---
    ext = {r.key for i in infos for r in i.tables if not r.internal}
    n_src = sum(1 for k in ext if lookup_entries(entries, k)[1])
    n_tgt = sum(1 for k in ext if any(match_entry(t, k) for t in tgt_keys))
    if n_tgt and not n_src and not force:
        return _fail(res, "Input already uses target-pod tables and none of the source tables; "
                          "not converting again (to avoid double conversion).")
    if n_tgt and n_src:
        res.warnings.append("Input mixes source-pod and target-pod table names; verify the result.")

    # --- 5-7: tables ---
    kept_schemas, unmapped_t, unmapped_c, pending = set(), set(), set(), set()
    names_by_schema = defaultdict(set)
    for i in infos:
        for r in i.tables:
            if not r.internal:
                names_by_schema[r.node.name.lower()].add(r.node.db.lower())
    for i in infos:
        for r in i.tables:
            if r.internal:
                continue
            ek, targets = lookup_entries(entries, r.key)
            distinct = sorted(set(targets))
            if not distinct:
                if r.key not in unmapped_t:
                    res.warnings.append(f"Unmapped table: {r.key} (left unchanged)")
                unmapped_t.add(r.key)
            elif len(distinct) > 1:
                S["ambiguous_mappings"] += 1
                res.warnings.append(f"Ambiguous table mapping for {r.key}: {', '.join(distinct)} (left unchanged)")
            else:
                if len(ek.split(".")) < len(r.key.split(".")) and len(names_by_schema[r.node.name.lower()]) > 1:
                    res.warnings.append(f"Table '{r.node.name}' appears under several schemas; unqualified mapping "
                                        f"entry applied to {r.key}. Verify.")
                if r.role in ("insert", "update", "delete", "merge", "drop"):
                    res.warnings.append(f"Statement {i.index}: {r.role.upper()} on external table {r.key}; verify target.")
                old = r.key
                _apply_table(r, distinct[0], schema_map, kept_schemas)
                r.pair = (ek, norm_key(distinct[0]))
                S["tables_mapped"] += 1
                res.applied.append(f"table {old} -> {'.'.join(p for p in r.renamed_to if p)}")
    for s in sorted(kept_schemas):
        res.warnings.append(f"Schema '{s}' kept from source: no target schema in table mapping or 'Schemas' sheet.")

    # --- 6-8: columns, only in the context of their own table mapping ---
    stars = set()
    for i in infos:
        for c in i.columns:
            if c.status in ("ambiguous", "unresolved"):
                S["ambiguous_mappings"] += 1
                res.warnings.append(f"Statement {i.index}: cannot tell which table column "
                                    f"'{c.node.sql(dialect=DIALECT)}' belongs to ({c.status}); left unchanged. "
                                    "Qualify it with a table alias.")
                continue
            ref = c.table
            if ref is None or ref.internal or ref.pair is None:
                continue
            if c.status == "star":
                stars.add(ref.key)
                continue
            changed = False
            if c.kind == "column":
                changed = _requalify(c.node, ref)
            cands = colmap.get(ref.pair, {}).get(c.name.lower(), [])
            confirmed = sorted({t for t, st in cands if st == "confirmed"})
            label = f"{ref.key}.{c.name}"
            if len(confirmed) == 1:
                new = confirmed[0]
                if new != c.name:
                    col = c.node
                    wrap = (preserve_internal_names and c.kind == "column" and isinstance(col.parent, exp.Select)
                            and col.arg_key == "expressions" and col.parent.find_ancestor(exp.CTE, exp.Subquery, exp.Create))
                    rename_node(col, new)
                    if wrap:
                        col.replace(exp.alias_(col.copy(), c.name))
                        S["aliases_added"] += 1
                    S["columns_mapped"] += 1
                    changed = True
                    res.applied.append(f"column {label} -> {new}")
            elif len(confirmed) > 1 or any(st == "ambiguous" for _, st in cands):
                S["ambiguous_mappings"] += 1
                res.warnings.append(f"Ambiguous column mapping for {label}: left unchanged.")
            elif cands:
                if label not in pending:
                    res.warnings.append(f"Column {label}: mapping exists but is not 'confirmed' ({cands[0][1]}); left unchanged.")
                pending.add(label)
            else:
                if label not in unmapped_c and d.column_rows:
                    res.warnings.append(f"No column mapping for {label} (table is mapped); left unchanged.")
                unmapped_c.add(label)
            if changed and c.kind == "column" and _in_join_condition(c.node):
                S["join_references_updated"] += 1
    for i in infos:
        for j in i.joins:
            if j["using"]:
                S["ambiguous_mappings"] += 1
                res.warnings.append(f"Statement {i.index}: JOIN ... USING ({', '.join(j['using'])}) not rewritten; "
                                    "verify column names exist in both tables.")
    for k in sorted(stars):
        res.warnings.append(f"SELECT * / alias.* used on {k}: columns now come from the target table; verify downstream use.")
    S["unmapped_objects"] = len(unmapped_t) + len(unmapped_c) + len(pending)

    # --- 9-10: generate, restore schemas, re-parse as a syntax check ---
    try:
        out = ";\n\n".join(s.sql(dialect=DIALECT, pretty=True) for s in stmts) + ";"
        sqlglot.parse(out, read=DIALECT)          # syntax check while schemas are still temp ids
        out = registry.restore(out)
    except Exception as e:  # noqa
        return _fail(res, f"Generated SQL failed validation: {e}")
    res.ok, res.sql = True, out
    res.warnings.append("Output is regenerated by SQLGlot: formatting is normalised and comments may move slightly.")
    return res
