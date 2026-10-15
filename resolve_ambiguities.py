#!/usr/bin/env python
"""Resolve ambiguous / suggested column mappings in pod_mapping.xlsx.

Interactive:   python resolve_ambiguities.py analysis_out/pod_mapping.xlsx
Via CSV:       python resolve_ambiguities.py pod_mapping.xlsx --export pending.csv      (fill 'target_column')
               python resolve_ambiguities.py pod_mapping.xlsx --decisions pending.csv
Options: --include-suggested (also review 'suggested' rows), --out reviewed.xlsx
Nothing is ever confirmed automatically; every confirmation here is a human decision.
"""
import argparse
import csv
import sys
from collections import defaultdict
from pathlib import Path

from mapping_book import MappingBook, write_mapping_workbook

REVIEW = {"ambiguous"}


def _key(r):
    return (r["source_table"].lower(), r["target_table"].lower(), r["source_column"].lower())


def _rank(r):
    try:
        return float(r.get("rank") or 1)
    except ValueError:
        return 1.0


def pending_groups(rows, include_suggested=False):
    """{key: [candidate rows]} for source columns that still need a human decision."""
    wanted = REVIEW | ({"suggested"} if include_suggested else set())
    groups = defaultdict(list)
    for r in rows:
        if r.get("source_column"):
            groups[_key(r)].append(r)
    return {k: sorted(v, key=_rank)
            for k, v in groups.items()
            if any((r.get("status") or "").lower() in wanted for r in v)
            and not any((r.get("status") or "").lower() == "confirmed" for r in v)}


def apply_choice(rows, key, target):
    """Confirm `target` for the group; drop the other candidates. Empty target = reject all."""
    group = [r for r in rows if r.get("source_column") and _key(r) == key]
    rest = [r for r in rows if not (r.get("source_column") and _key(r) == key)]
    if not group:
        raise KeyError(key)
    base = dict(group[0])
    if not target:
        base.update(target_column="", status="unmapped", match_method="manual", confidence=0,
                    notes="all candidates rejected by reviewer")
        return rest + [base]
    hit = next((r for r in group if r["target_column"].lower() == target.lower()), None)
    row = dict(hit or base)
    row.update(target_column=hit["target_column"] if hit else target, status="confirmed",
               notes="confirmed manually" if hit else "manual entry by reviewer")
    if not hit:
        row.update(match_method="manual", confidence=1.0)
    return rest + [row]


def conflicts(rows):
    """Confirmed rows where two source columns of one table pair share a target column."""
    seen = defaultdict(list)
    for r in rows:
        if (r.get("status") or "").lower() == "confirmed" and r.get("target_column"):
            seen[(r["source_table"].lower(), r["target_table"].lower(), r["target_column"].lower())].append(r["source_column"])
    return {k: v for k, v in seen.items() if len({x.lower() for x in v}) > 1}


def _show(cands):
    print(f"\n{cands[0]['source_table']} -> {cands[0]['target_table']}   source column: {cands[0]['source_column']}")
    shown = [c for c in cands if c["target_column"]]
    for i, c in enumerate(shown, 1):
        print(f"  [{i}] {c['target_column']:<30} {c.get('match_method', ''):<17} "
              f"conf={c.get('confidence', '')}  {c.get('notes', '')}")
    return shown


def interactive(rows, include_suggested):
    groups = pending_groups(rows, include_suggested)
    print(f"{len(groups)} source column(s) need a decision.  "
          "Enter number = accept, t = type a target name, n = no match, s = skip, q = save and quit")
    for key, cands in groups.items():
        shown = _show(cands)
        while True:
            a = input("  choice> ").strip().lower()
            if a == "q":
                return rows
            if a == "s":
                break
            if a == "n":
                rows = apply_choice(rows, key, "")
                break
            if a == "t":
                t = input("  target column name> ").strip()
                if t:
                    rows = apply_choice(rows, key, t)
                    break
            elif a.isdigit() and 1 <= int(a) <= len(shown):
                rows = apply_choice(rows, key, shown[int(a) - 1]["target_column"])
                if any(key[0] == k[0] and key[1] == k[1] for k in conflicts(rows)):
                    print("  WARNING: another source column in this table pair is also confirmed to the same target.")
                break
            else:
                print("  not understood")
    return rows


def export_csv(rows, path, include_suggested):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["source_table", "target_table", "source_column", "candidates", "target_column"])
        for cands in pending_groups(rows, include_suggested).values():
            w.writerow([cands[0]["source_table"], cands[0]["target_table"], cands[0]["source_column"],
                        "; ".join(f"{c['target_column']} ({c.get('confidence')})" for c in cands if c["target_column"]), ""])
    print(f"Wrote {path}. Fill 'target_column' (blank = reject all candidates), then run with --decisions.")


def apply_csv(rows, path):
    n = 0
    with open(path, newline="", encoding="utf-8") as f:
        for d in csv.DictReader(f):
            key = (d["source_table"].lower(), d["target_table"].lower(), d["source_column"].lower())
            if not any(r.get("source_column") and _key(r) == key for r in rows):
                continue
            rows = apply_choice(rows, key, (d.get("target_column") or "").strip())
            n += 1
    print(f"Applied {n} decision(s).")
    return rows


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--out")
    ap.add_argument("--export")
    ap.add_argument("--decisions")
    ap.add_argument("--include-suggested", action="store_true")
    a = ap.parse_args(argv)
    book = MappingBook.load(a.workbook)
    rows = book.column_rows
    if a.export:
        return export_csv(rows, a.export, a.include_suggested)
    rows = apply_csv(rows, a.decisions) if a.decisions else interactive(rows, a.include_suggested)
    for (s, t, tc), srcs in conflicts(rows).items():
        print(f"WARNING: {s} -> {t}: target column '{tc}' confirmed for several source columns: {', '.join(srcs)}")
    left = len(pending_groups(rows, a.include_suggested))
    out = a.out or str(Path(a.workbook).with_name(Path(a.workbook).stem + ".reviewed.xlsx"))
    write_mapping_workbook(out, book.labels, book.table_rows, rows, book.schema_rows)
    print(f"Saved {out}  ({left} item(s) still pending).")


if __name__ == "__main__":
    sys.exit(main())
