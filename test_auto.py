import tempfile
import unittest
from pathlib import Path

import pandas as pd

import auto


class ParseTests(unittest.TestCase):
    def test_multiple_records(self):
        lines = ["name: Ali", "city: Kashan", "", "", "name: Sara", "city: Tehran"]
        self.assertEqual(auto.parse_records(lines), [
            {"name": "Ali", "city": "Kashan"},
            {"name": "Sara", "city": "Tehran"},
        ])

    def test_colon_in_value_is_kept(self):
        self.assertEqual(auto.parse_records(["time: 10:30"]), [{"time": "10:30"}])

    def test_malformed_lines_are_skipped(self):
        with self.assertLogs("auto", level="WARNING"):
            result = auto.parse_records(["name: Ali", "garbage line", ": no key"])
        self.assertEqual(result, [{"name": "Ali"}])

    def test_bom_is_removed(self):
        self.assertEqual(auto.parse_records(["\ufeffname: Ali"]), [{"name": "Ali"}])

    def test_empty_input(self):
        self.assertEqual(auto.parse_records([]), [])


class CliTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def test_end_to_end(self):
        src, out = self.tmp / "in.txt", self.tmp / "out.xlsx"
        src.write_text("name: Ali\ncity: Kashan\n\nname: Sara\n", encoding="utf-8")
        self.assertEqual(auto.main([str(src), str(out), "--sheet", "People"]), 0)
        df = pd.read_excel(out, sheet_name="People")
        self.assertEqual(list(df.columns), ["name", "city"])
        self.assertEqual(len(df), 2)

    def test_missing_input_returns_error(self):
        self.assertEqual(auto.main([str(self.tmp / "nope.txt"), str(self.tmp / "o.xlsx")]), 1)

    def test_invalid_sheet_name(self):
        src = self.tmp / "in.txt"
        src.write_text("a: 1\n", encoding="utf-8")
        self.assertEqual(auto.main([str(src), str(self.tmp / "o.xlsx"), "--sheet", "bad/name"]), 1)


if __name__ == "__main__":
    unittest.main()
