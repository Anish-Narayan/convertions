"""Name-based column candidate matching. Produces CANDIDATES only - never claims semantic equivalence."""
import json
import re
from collections import defaultdict
from difflib import SequenceMatcher

DEFAULT_ABBREVIATIONS = {"amt": "amount", "dt": "date", "dte": "date", "num": "number", "nbr": "number",
                         "qty": "quantity", "cd": "code", "desc": "description", "descr": "description",
                         "flg": "flag"}


class ColumnMatcher:
    def __init__(self, abbreviations=None, synonym_groups=None, min_score=0.5, margin=0.08, top_n=3):
        self.abbr = dict(DEFAULT_ABBREVIATIONS)
        self.abbr.update({k.lower(): v.lower() for k, v in (abbreviations or {}).items()})
        # synonym_groups: user-supplied [["member","beneficiary"], ...]; none are assumed by default
        self.syn = {t.lower(): g[0].lower() for g in (synonym_groups or []) for t in g}
        self.min_score, self.margin, self.top_n = min_score, margin, top_n

    @classmethod
    def from_files(cls, abbreviations_json=None, synonyms_json=None, **kw):
        load = lambda p: json.load(open(p, encoding="utf-8")) if p else None
        return cls(load(abbreviations_json), load(synonyms_json), **kw)

    def tokens(self, name):
        s = re.sub(r"([a-z0-9])([A-Z])", r"\1_\2", name)
        toks = [t for t in re.split(r"[^A-Za-z0-9]+", s.lower()) if t]
        return [self.syn.get(self.abbr.get(t, t), self.abbr.get(t, t)) for t in toks]

    def score(self, src, tgt):
        if src == tgt:
            return "exact", 1.0
        if src.lower() == tgt.lower():
            return "case_insensitive", 0.98
        a, b = self.tokens(src), self.tokens(tgt)
        if a == b:
            return "normalized", 0.95
        sa, sb = set(a), set(b)
        jac = len(sa & sb) / len(sa | sb) if sa | sb else 0
        seq = SequenceMatcher(None, " ".join(sorted(a)), " ".join(sorted(b))).ratio()
        return "token_fuzzy", round(min(0.9, 0.5 * jac + 0.5 * seq), 3)

    def match(self, src_cols, tgt_cols):
        rows, top = [], {}
        for s in src_cols:
            scored = sorted(((t,) + self.score(s, t) for t in tgt_cols), key=lambda x: (-x[2], x[0]))
            cands = [c for c in scored if c[2] >= self.min_score][: self.top_n]
            if not cands:
                rows.append(dict(source_column=s, target_column="", rank=1, match_method="none",
                                 confidence=0.0, status="unmapped", notes="no candidate above threshold"))
                continue
            close = len(cands) > 1 and cands[0][2] - cands[1][2] <= self.margin and cands[0][1] != "exact"
            for rank, (t, method, sc) in enumerate(cands, 1):
                if cands[0][1] == "exact" and rank == 1:
                    status = "confirmed"
                elif close:
                    status = "ambiguous"
                else:
                    status = "suggested"
                rows.append(dict(source_column=s, target_column=t, rank=rank, match_method=method,
                                 confidence=sc, status=status,
                                 notes="name similarity only; verify meaning" if status != "confirmed" else "identical name"))
            top[s] = cands[0][0]
        # conflict: several source columns claim the same target column
        claims = defaultdict(list)
        for s, t in top.items():
            claims[t].append(s)
        for t, srcs in claims.items():
            if len(srcs) > 1:
                for r in rows:
                    if r["target_column"] == t and r["rank"] == 1 and r["match_method"] != "exact" and r["source_column"] in srcs:
                        r["status"] = "ambiguous"
                        r["notes"] = f"target also claimed by: {', '.join(x for x in srcs if x != r['source_column'])}"
        matched = {r["target_column"] for r in rows if r["rank"] == 1 and r["target_column"]}
        for t in tgt_cols:
            if t not in matched:
                rows.append(dict(source_column="", target_column=t, rank=1, match_method="none", confidence=0.0,
                                 status="unmapped", notes="target column has no source counterpart"))
        return rows
