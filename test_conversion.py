"""End-to-end conversion tests (require sqlglot). Skipped automatically when it is not installed."""
import os, sys, tempfile, unittest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
try:
    import sqlglot  # noqa
    HAVE = True
except ImportError:
    HAVE = False
if HAVE:
    from converter import convert_sql
    from mapping_book import MappingBook, write_mapping_workbook

COLS = [("claims", "claim_header", "member_id", "beneficiary_id"), ("claims", "claim_header", "claim_amount", "allowed_amount"),
        ("claims", "claim_header", "provider_id", "provider_key"), ("claims", "claim_header", "service_date", "date_of_service"),
        ("providers", "provider_master", "member_id", "contact_id"), ("providers", "provider_master", "provider_id", "provider_key")]


@unittest.skipUnless(HAVE, "sqlglot not installed")
class T(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.d = tempfile.mkdtemp(); p = os.path.join(cls.d, "m.xlsx")
        write_mapping_workbook(p, ("PodA", "PodB"), [("claims", "claim_header"), ("providers", "provider_master")],
                               [dict(source_table=a, target_table=b, source_column=c, target_column=t, status="confirmed")
                                for a, b, c, t in COLS])
        cls.book = MappingBook.load(p)

    def conv(self, sql):
        return convert_sql(sql, self.book, "PodA", "PodB")

    def test_alias_join_same_named_columns_literals_schema(self):
        r = self.conv("SELECT a.member_id, p.member_id AS pm FROM $(gold.schema).claims a "
                      "LEFT JOIN $(gold.schema).providers p ON a.provider_id = p.provider_id "
                      "WHERE a.service_date >= '2024-01-01' AND a.member_id <> 'member_id'")
        self.assertTrue(r.ok, r.report())
        o = r.sql
        self.assertIn("$(gold.schema).claim_header", o); self.assertIn("a.beneficiary_id", o)
        self.assertIn("p.contact_id AS pm", o)          # same column name, different table -> different target
        self.assertIn("'member_id'", o)                 # literal untouched
        self.assertIn("LEFT JOIN", o); self.assertEqual(r.summary["join_references_updated"], 2)

    def test_cte_keeps_internal_name(self):
        r = self.conv("WITH c AS (SELECT member_id FROM claims) SELECT member_id FROM c")
        self.assertTrue(r.ok, r.report())
        self.assertIn("beneficiary_id AS member_id", r.sql); self.assertIn("FROM c", r.sql)

    def test_created_table_not_mapped(self):
        r = self.conv("DROP TABLE IF EXISTS tmp_x; CREATE TABLE tmp_x AS SELECT member_id FROM claims; "
                      "INSERT INTO tmp_x SELECT member_id FROM claims; SELECT member_id FROM tmp_x;")
        self.assertTrue(r.ok, r.report())
        self.assertIn("tmp_x", r.sql); self.assertEqual(r.sql.count("claim_header"), 2)
        self.assertEqual(r.summary["tables_mapped"], 2)

    def test_nested_subquery_and_correlated(self):
        r = self.conv("SELECT x.member_id FROM (SELECT member_id FROM claims) x "
                      "WHERE EXISTS (SELECT 1 FROM providers p WHERE p.provider_id = x.member_id)")
        self.assertTrue(r.ok, r.report()); self.assertIn("provider_key", r.sql)

    def test_ambiguous_unqualified_column_not_guessed(self):
        r = self.conv("SELECT member_id FROM claims a JOIN providers p ON a.provider_id = p.provider_id")
        self.assertTrue(r.ok); self.assertIn("member_id", r.sql)
        self.assertGreaterEqual(r.summary["ambiguous_mappings"], 1)

    def test_unmapped_table_and_parse_error(self):
        r = self.conv("SELECT * FROM other_table"); self.assertEqual(r.summary["unmapped_objects"], 1)
        r = self.conv("SELEC FROM FROM"); self.assertFalse(r.ok); self.assertEqual(r.sql, "")

    def test_no_double_conversion(self):
        first = self.conv("SELECT member_id FROM claims").sql
        r = self.conv(first); self.assertFalse(r.ok)
        self.assertTrue(convert_sql(first, self.book, "PodA", "PodB", force=True).ok)

    def test_reverse_direction(self):
        r = convert_sql("SELECT beneficiary_id FROM claim_header", self.book, "PodB", "PodA")
        self.assertTrue(r.ok, r.report()); self.assertIn("member_id", r.sql)


if __name__ == "__main__":
    unittest.main()
