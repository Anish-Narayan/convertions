"""Tests that need no SQLGlot: schema registry, column matching, workbook round trip."""
import os, sys, tempfile, unittest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from sql_common import SchemaRegistry, norm_key, match_entry, lookup_entries
from column_matching import ColumnMatcher
from mapping_book import MappingBook, write_mapping_workbook


class T(unittest.TestCase):
    def test_registry_consistent_and_distinct(self):
        r = SchemaRegistry()
        a = r.replace("SELECT * FROM $(gold.schema).claims JOIN $(silver.schema).m")
        b = r.replace("SELECT * FROM $(gold.schema).x")
        self.assertEqual(a, "SELECT * FROM temp1.claims JOIN temp2.m")
        self.assertEqual(b, "SELECT * FROM temp1.x")
        self.assertEqual(r.restore(a), "SELECT * FROM $(gold.schema).claims JOIN $(silver.schema).m")

    def test_registry_avoids_existing_temp_ids(self):
        r = SchemaRegistry()
        r.reserve("SELECT 1 FROM temp1.t")
        self.assertEqual(r.replace("$(a.b).t"), "temp2.t")
        r2 = SchemaRegistry(); r2.replace("$(a)")
        with self.assertRaises(ValueError):
            r2.replace("SELECT * FROM temp1.t")

    def test_name_matching(self):
        self.assertTrue(match_entry("claims", "a.claims"))
        self.assertFalse(match_entry("a.claims", "b.claims"))
        self.assertFalse(match_entry("a.claims", "claims"))
        self.assertEqual(lookup_entries({"claims": [1], "a.claims": [2]}, "a.claims")[1], [2])

    def test_column_matching(self):
        m = ColumnMatcher()
        rows = m.match(["member_id", "claim_amt", "svc"], ["member_id", "ClaimAmount", "other"])
        by = {r["source_column"]: r for r in rows if r["source_column"]}
        self.assertEqual(by["member_id"]["status"], "confirmed")
        self.assertEqual(by["claim_amt"]["target_column"], "ClaimAmount")
        self.assertIn(by["claim_amt"]["status"], ("suggested",))
        self.assertEqual(by["svc"]["status"], "unmapped")

    def test_synonyms_only_when_supplied(self):
        self.assertEqual(ColumnMatcher().match(["member_id"], ["beneficiary_id"])[0]["status"], "unmapped")
        m = ColumnMatcher(synonym_groups=[["member", "beneficiary"]])
        r = m.match(["member_id"], ["beneficiary_id"])[0]
        self.assertEqual((r["match_method"], r["status"]), ("normalized", "suggested"))

    def test_workbook_roundtrip_and_direction(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "m.xlsx")
            write_mapping_workbook(p, ("PodA", "PodB"), [("claims", "claim_header")],
                                   [dict(source_table="claims", target_table="claim_header", source_column="a",
                                         target_column="b", status="confirmed")], [("$(x)", "$(y)")])
            b = MappingBook.load(p)
            self.assertEqual(b.labels, ("PodA", "PodB"))
            rev = b.directed("PodB", "PodA")
            self.assertEqual(rev.table_rows, [("claim_header", "claims")])
            self.assertEqual(rev.column_rows[0]["source_column"], "b")
            with self.assertRaises(ValueError):
                b.directed("PodA", "PodA")


if __name__ == "__main__":
    unittest.main()
