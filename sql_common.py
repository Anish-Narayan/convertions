"""Shared helpers: dialect, concept splitting (from the legacy script), schema placeholders, name matching."""
import re
from pathlib import Path

DIALECT = "databricks"
PLACEHOLDER = re.compile(r"\$\(([^()\s]+)\)")                      # $(gold.schema)
HEADER = re.compile(r"^\s*--([A-Za-z0-9_]+)\s*$", re.MULTILINE)    # legacy: --concept_name
_TEMP = re.compile(r"\btemp(\d+)\b")


def read_sql(path):
    return Path(path).read_text(encoding="utf-8-sig")


def split_concepts(text, require_headers=False):
    """Legacy logic. Differences: a file without headers is one concept unless require_headers=True."""
    matches = list(HEADER.finditer(text))
    if not matches:
        if require_headers:
            raise ValueError("No concept headers found, e.g. --wrong_medication")
        return [{"name": "(whole_file)", "sql": text.strip()}]
    out = []
    for i, m in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        out.append({"name": m.group(1), "sql": text[m.end():end].strip()})
    return out


class SchemaRegistry:
    """placeholder <-> tempN. Share ONE instance across files so ids stay consistent."""

    def __init__(self):
        self.placeholder_to_temp, self.temp_to_placeholder = {}, {}
        self._reserved, self._counter = set(), 0

    def reserve(self, *texts):
        """Note tempN tokens that already exist as real identifiers so they are never allocated."""
        for text in texts:
            for n in _TEMP.findall(text):
                if f"temp{n}" in self.temp_to_placeholder:
                    raise ValueError(f"Input already contains identifier 'temp{n}', which clashes "
                                     "with a temporary schema id. Rename it and retry.")
                self._reserved.add(int(n))

    def _allocate(self):
        while True:
            self._counter += 1
            if self._counter not in self._reserved:
                return f"temp{self._counter}"

    def replace(self, text):
        self.reserve(text)

        def sub(m):
            ph = m.group(0)
            if ph not in self.placeholder_to_temp:
                t = self._allocate()
                self.placeholder_to_temp[ph], self.temp_to_placeholder[t] = t, ph
            return self.placeholder_to_temp[ph]
        return PLACEHOLDER.sub(sub, text)

    def restore(self, text):
        return _TEMP.sub(lambda m: self.temp_to_placeholder.get(m.group(0), m.group(0)), text)


# ---- table-name helpers -------------------------------------------------------------------
def table_key(t):
    """Lower-cased fully qualified key of a sqlglot Table node."""
    return ".".join(p.lower() for p in (t.catalog, t.db, t.name) if p)


def split_parts(name):
    return [p for p in (x.strip().strip('`"') for x in name.strip().split(".")) if p]


def norm_key(name):
    return ".".join(p.lower() for p in split_parts(name))


def match_entry(entry_key, sql_key):
    """Entry 'claims' matches 'a.claims'; entry 'a.claims' never matches 'b.claims' or bare 'claims'."""
    e, s = entry_key.split("."), sql_key.split(".")
    return len(e) <= len(s) and s[-len(e):] == e


def lookup_entries(entries, sql_key):
    """Most specific entry first. Returns (entry_key, value) or (None, [])."""
    parts = sql_key.split(".")
    for n in range(len(parts), 0, -1):
        k = ".".join(parts[-n:])
        if k in entries:
            return k, entries[k]
    return None, []
