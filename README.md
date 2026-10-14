# SQL Pod Cross-Pollination Tool

## Layout
| File | Role |
|---|---|
| `sql_common.py` | dialect, legacy `split_concepts`, `SchemaRegistry` (`$(x)` <-> `tempN`), table-name matching |
| `sql_resolver.py` | SQLGlot scope-based table/column resolution (shared by A and B), legacy `fingerprint` |
| `column_matching.py` | name-based column candidates (exact / case / normalised / fuzzy) |
| `mapping_book.py` | read/write the Excel mapping workbook (openpyxl) |
| `analysis_engine.py` | **Component A** (CLI) |
| `converter.py`, `converter_gui.py` | **Component B** engine + Tkinter GUI (the .exe) |
| `tests/` | `test_pure.py` (runs anywhere), `test_conversion.py` (needs sqlglot) |

## Install and test
```
python -m venv .venv && .venv\Scripts\activate        (Linux/macOS: source .venv/bin/activate)
pip install -r requirements-analysis.txt
python -m unittest discover tests -v
```

## Component A
```
python analysis_engine.py source.sql target.sql table_mapping.xlsx --out analysis_out
   [--synonyms synonyms.json] [--abbreviations abbr.json] [--min-score 0.5] [--no-similarity]
```
`table_mapping.xlsx`: row 1 = pod labels (A1 = source pod, B1 = target pod), then `source_table | target_table`.
Write schema placeholders in the Excel exactly as in the SQL, e.g. `$(gold.schema).claims`.

Outputs in `--out`: `source_metadata.json`, `target_metadata.json`, `mapping_review.xlsx` (TableMapping, ColumnMapping,
Log, StructuralSimilarity), `pod_mapping.xlsx` (reusable), `analysis_log.txt`, `source.temp.sql`, `target.temp.sql`,
`schema_map.json`. Originals are never touched.

**Review step:** in `pod_mapping.xlsx` > `Columns`, change `status` to `confirmed` for candidates you accept (and
delete the rejected alternatives). The converter applies **only `confirmed`** rows. Only identical column names are
auto-confirmed. Fill the optional `Schemas` sheet (pod-A placeholder -> pod-B placeholder) to remap schemas.
Synonyms such as member/beneficiary are never assumed; supply them with `--synonyms` if you want them considered.

## Component B (GUI)
```
python converter_gui.py
```
Upload `pod_mapping.xlsx`; dropdowns are filled from the labels in `Tables!A1:B1`; both directions work
(reverse uses the inverted mappings).

## Build the .exe (Windows)
```
build_exe.bat
```
which runs: `pyinstaller --noconfirm --onefile --windowed --name SQLConverter --collect-all sqlglot converter_gui.py`
Result: `dist\SQLConverter.exe` (no Python needed on the user machine; ship it with the mapping workbook).
