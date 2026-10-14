#!/usr/bin/env python
"""Component A - SQL analysis and mapping engine.

python analysis_engine.py SOURCE.sql TARGET.sql TABLE_MAPPING.xlsx --out out_dir
Originals are never modified. Temporary (placeholder-free) copies are written to the output folder.
"""
import argparse
import json
import sys
from collections import defaultdict
from itertools import combinations
from pathlib import Path

import sqlglot

from column_matching import ColumnMatcher
from mapping_book import MappingBook, write_mapping_workbook, write_report_workbook
from sql_common import (DIALECT, SchemaRegistry, lookup_entries, match_entry, norm_key,
                        read_sql, split_concepts)
from sql_resolver import fingerprint, resolve_script

THRESHOLD = 0.65   # legacy structural-similarity threshold


def split_statements(sql):
    """Quote/comment-aware ';' splitter, used only as a fallback to isolate a bad statement."""
    out, buf, i, n, quote = [], [], 0, len(sql), None
    while i < n:
        ch = sql[i]
        if quote:
            buf.append(ch)
            if ch == "\\" and i + 1 < n:
                buf.append(sql[i + 1]); i += 2; continue
            if ch == quote:
                quote = None
        elif ch in "'\"`":
            quote = ch; buf.append(ch)
        elif sql.startswith("--", i):
            j = sql.find("\n", i); j = n if j < 0 else j
            buf.append(sql[i:j]); i = j; continue
        elif sql.startswith("/*", i):
            j = sql.find("*/", i + 2); j = n if j < 0 else j + 2
            buf.append(sql[i:j]); i = j; continue
        elif ch == ";":
            out.append("".join(buf)); buf = []
        else:
            buf.append(ch)
        i += 1
    out.append("".join(buf))
    return [s for s in out if s.strip()]


def parse_statements(sql):
    try:
        return [s for s in sqlglot.parse(sql, read=DIALECT) if s is not None], []
    except Exception as e:  # noqa
        stmts, errs = [], [f"whole-concept parse failed, retrying per statement: {e}"]
        for i, part in enumerate(split_statements(sql), 1):
            try:
                stmts += [s for s in sqlglot.parse(part, read=DIALECT) if s is not None]
            except Exception as e2:  # noqa
                errs.append(f"statement {i} skipped: {e2}")
        return stmts, errs


def _display(node):
    return ".".join(p for p in (node.catalog, node.db, node.name) if p)


def analyze_script(label, path, tmp_text, registry):
    meta = {"label": label, "file": str(path), "dialect": DIALECT,
            "schemas": dict(registry.placeholder_to_temp), "concepts": [], "statements": [],
            "external_tables": {}, "internal_tables": {}, "ctes": [], "joins": [],
            "dependencies": [], "review_columns": [], "parse_errors": []}
    ext, sim_docs = {}, []
    for c in split_concepts(tmp_text):
        stmts, errs = parse_statements(c["sql"])
        meta["parse_errors"] += [f"[{c['name']}] {e}" for e in errs]
        infos, state = resolve_script(stmts)
        meta["concepts"].append({"name": c["name"], "statement_count": len(stmts)})
        sim_docs.append({"name": c["name"], "fingerprint": " ".join(fingerprint(s) for s in stmts)
                         + " " + " ".join(j["type"] for i in infos for j in i.joins)
                         + " " + " ".join(o for i in infos for o in i.operations)})
        for info in infos:
            meta["parse_errors"] += [f"[{c['name']}] stmt {info.index}: {e}" for e in info.errors]
            reads = [r for r in info.tables if r.role == "read"]
            meta["statements"].append({
                "concept": c["name"], "index": info.index, "kind": info.kind,
                "operations": info.operations, "created_table": registry.restore(info.created) if info.created else None,
                "reads": sorted({registry.restore(r.key) for r in reads}),
                "writes": sorted({registry.restore(r.key) for r in info.tables if r.role != "read"}),
                "subqueries": info.subqueries, "warnings": info.warnings,
                "column_aliases": info.aliases})
            meta["ctes"] += [{"concept": c["name"], "statement": info.index, "name": n} for n in info.ctes]
            for j in info.joins:
                meta["joins"].append({"concept": c["name"], "statement": info.index, **j})
            for r in info.tables:
                if r.internal:
                    it = meta["internal_tables"].setdefault(registry.restore(r.key), {
                        "concept": c["name"], "created_in_statement": state.created.get(r.key),
                        "defined_columns": state.created_columns.get(r.key, []), "used_in_statements": []})
                    if r.role != "create" and info.index not in it["used_in_statements"]:
                        it["used_in_statements"].append(info.index)
                    prod = state.created.get(r.key)
                    if prod and prod != info.index:
                        meta["dependencies"].append({"concept": c["name"], "from_statement": prod,
                                                     "to_statement": info.index, "table": registry.restore(r.key)})
                    continue
                e = ext.setdefault(r.key, {"table": registry.restore(_display(r.node)), "key": r.key,
                                           "aliases": set(), "roles": set(), "columns": defaultdict(int),
                                           "uses_star": False, "concepts": set()})
                e["roles"].add(r.role); e["concepts"].add(c["name"])
                if r.alias:
                    e["aliases"].add(r.alias)
            for col in info.columns:
                if col.table is not None and not col.table.internal:
                    e = ext[col.table.key]
                    if col.status == "star":
                        e["uses_star"] = True
                    elif col.status == "resolved":
                        e["columns"][col.name] += 1
                elif col.status in ("ambiguous", "unresolved") and col.node is not None:
                    meta["review_columns"].append({
                        "concept": c["name"], "statement": info.index, "column": registry.restore(col.node.sql(dialect=DIALECT)),
                        "reason": "source table cannot be resolved unambiguously (flagged, not guessed)"
                        if col.status == "ambiguous" else "qualifier/source not found in scope"})
    meta["external_tables"] = {k: {**v, "aliases": sorted(v["aliases"]), "roles": sorted(v["roles"]),
                                   "concepts": sorted(v["concepts"]), "columns": dict(sorted(v["columns"].items()))}
                               for k, v in ext.items()}
    return meta, sim_docs


def build_table_report(book, src, tgt, registry):
    entries = defaultdict(list)       # source entry key -> [(raw_s, raw_t, t_key)]
    for s_raw, t_raw in book.table_rows:
        entries[norm_key(registry.replace(s_raw))].append((s_raw, t_raw, norm_key(registry.replace(t_raw))))
    rows, pairs, log, used_entries = [], [], [], set()
    for s_key, s in src["external_tables"].items():
        ek, hits = lookup_entries(entries, s_key)
        if not hits:
            rows.append([s["table"], "", "", "", "unmapped_source_table", "not in supplied mapping"])
            log.append(f"UNMAPPED source table: {s['table']}")
            continue
        used_entries.add(ek)
        tkeys = sorted({h[2] for h in hits})
        found = sorted({k for tk in tkeys for k in tgt["external_tables"] if match_entry(tk, k)})
        if len(tkeys) > 1:
            status, note = "ambiguous", "source table maps to several target tables"
        elif not found:
            status, note = "target_not_in_sql", "mapped, but target table not referenced in target SQL"
        elif len(found) > 1:
            status, note = "ambiguous", "unqualified target entry matches several schemas in target SQL"
        else:
            status, note = "mapped", ""
            pairs.append((s_key, found[0], hits[0][0], hits[0][1]))
        if status != "mapped":
            log.append(f"{status.upper()}: {s['table']} -> {', '.join(h[1] for h in hits)} ({note})")
        rows.append([s["table"], ", ".join(h[1] for h in hits),
                     ", ".join(tgt["external_tables"][k]["table"] for k in found), ek, status, note])
    for ek, hits in entries.items():
        if ek not in used_entries:
            rows.append(["", "", "", ek, "mapping_entry_not_in_source_sql", f"{hits[0][0]} -> {hits[0][1]}"])
    mapped_t = {p[1] for p in pairs}
    for k, t in tgt["external_tables"].items():
        if k not in mapped_t:
            rows.append(["", "", t["table"], "", "unmapped_target_table", "target table has no mapped source"])
            log.append(f"UNMAPPED target table: {t['table']}")
    return rows, pairs, log


def structural_similarity(src_docs, tgt_docs):
    """Legacy feature: TF-IDF cosine similarity of anonymised SQL structure (needs scikit-learn)."""
    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity
    except ImportError:
        return None
    docs = [d["fingerprint"] or "empty" for d in src_docs + tgt_docs]
    scores = cosine_similarity(TfidfVectorizer(ngram_range=(1, 3)).fit_transform(docs))
    ns = len(src_docs)
    return [[s["name"], t["name"], round(float(scores[i, ns + j]) * 100, 2),
             "similar" if scores[i, ns + j] >= THRESHOLD else ""]
            for i, s in enumerate(src_docs) for j, t in enumerate(tgt_docs)]


def run_analysis(source_sql, target_sql, mapping_xlsx, out_dir, matcher=None, similarity=True):
    out = Path(out_dir); out.mkdir(parents=True, exist_ok=True)
    book = MappingBook.load(mapping_xlsx)
    s_text, t_text = read_sql(source_sql), read_sql(target_sql)
    registry = SchemaRegistry()
    registry.reserve(s_text, t_text)                 # same registry => same placeholder -> same tempN in both
    s_tmp, t_tmp = registry.replace(s_text), registry.replace(t_text)
    (out / "source.temp.sql").write_text(s_tmp, encoding="utf-8")
    (out / "target.temp.sql").write_text(t_tmp, encoding="utf-8")
    (out / "schema_map.json").write_text(json.dumps(registry.placeholder_to_temp, indent=2), encoding="utf-8")

    src, s_docs = analyze_script(book.labels[0], source_sql, s_tmp, registry)
    tgt, t_docs = analyze_script(book.labels[1], target_sql, t_tmp, registry)
    (out / "source_metadata.json").write_text(json.dumps(src, indent=2, default=str), encoding="utf-8")
    (out / "target_metadata.json").write_text(json.dumps(tgt, indent=2, default=str), encoding="utf-8")

    t_rows, pairs, log = build_table_report(book, src, tgt, registry)
    matcher = matcher or ColumnMatcher()
    c_rows, wb_cols = [], []
    for s_key, t_key, s_raw, t_raw in pairs:
        s, t = src["external_tables"][s_key], tgt["external_tables"][t_key]
        if s["uses_star"] or t["uses_star"]:
            log.append(f"NOTE: SELECT * used for {s['table']} or {t['table']}; column list is incomplete")
        for m in matcher.match(sorted(s["columns"]), sorted(t["columns"])):
            c_rows.append([s["table"], t["table"], m["source_column"], m["target_column"], m["rank"],
                           m["match_method"], m["confidence"], m["status"], m["notes"]])
            if m["status"] == "unmapped":
                log.append(f"UNMAPPED column: {s['table']}.{m['source_column'] or '-'} / {t['table']}.{m['target_column'] or '-'}")
            elif m["status"] == "ambiguous":
                log.append(f"AMBIGUOUS column: {s['table']}.{m['source_column']} -> {m['target_column']} ({m['confidence']})")
            if m["source_column"] and m["target_column"]:
                wb_cols.append({"source_table": s_raw, "target_table": t_raw, **{k: m[k] for k in
                                ("source_column", "target_column", "status", "match_method", "confidence", "notes")}})
    for meta in (src, tgt):
        log += [f"PARSE [{meta['label']}]: {e}" for e in meta["parse_errors"]]
        log += [f"REVIEW [{meta['label']}] concept {r['concept']} stmt {r['statement']}: {r['column']} - {r['reason']}"
                for r in meta["review_columns"]]
    sheets = {"TableMapping": (["source_table", "mapped_target_entry", "target_table_in_sql", "matched_entry", "status", "notes"], t_rows),
              "ColumnMapping": (["source_table", "target_table", "source_column", "suggested_target_column", "rank",
                                 "match_method", "confidence", "status", "notes"], c_rows),
              "Log": (["message"], [[m] for m in log])}
    sim = structural_similarity(s_docs, t_docs) if similarity else None
    if sim:
        sheets["StructuralSimilarity"] = (["source_concept", "target_concept", "similarity_%", "flag"], sim)
    write_report_workbook(out / "mapping_review.xlsx", sheets)
    write_mapping_workbook(out / "pod_mapping.xlsx", book.labels, book.table_rows, wb_cols)
    (out / "analysis_log.txt").write_text("\n".join(log), encoding="utf-8")
    return {"pairs": len(pairs), "column_rows": len(c_rows), "log_entries": len(log), "out": str(out)}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("source_sql"); ap.add_argument("target_sql"); ap.add_argument("table_mapping_xlsx")
    ap.add_argument("--out", default="analysis_out")
    ap.add_argument("--abbreviations", help="JSON {abbr: expansion}")
    ap.add_argument("--synonyms", help='JSON [["member","beneficiary"], ...] (user-supplied; none assumed)')
    ap.add_argument("--min-score", type=float, default=0.5)
    ap.add_argument("--no-similarity", action="store_true")
    a = ap.parse_args(argv)
    m = ColumnMatcher.from_files(a.abbreviations, a.synonyms, min_score=a.min_score)
    print(json.dumps(run_analysis(a.source_sql, a.target_sql, a.table_mapping_xlsx, a.out, m, not a.no_similarity), indent=2))


if __name__ == "__main__":
    sys.exit(main())
