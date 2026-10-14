"""Tests for the fast xlsx writer (needs xlsxwriter + openpyxl, no database)."""
import os
import tempfile
import unittest

import numpy as np
import pandas as pd

from compare_app import xlsx_writer


class WriterTests(unittest.TestCase):
    def roundtrip(self, sheets):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, 'out.xlsx')
            xlsx_writer.write_workbook(path, sheets)
            return {name: pd.read_excel(path, sheet_name=name, engine='openpyxl', dtype=object)
                    for name in sheets}

    def test_values_and_headers_survive(self):
        df = pd.DataFrame({'Name': ['A One', None], 'Phone': ['9876543210', ''], 'Excel Row': [2, 3]})
        back = self.roundtrip({'S': df})['S']
        self.assertEqual(list(back.columns), ['Name', 'Phone', 'Excel Row'])
        self.assertEqual(back['Name'].iloc[0], 'A One')
        self.assertTrue(pd.isna(back['Name'].iloc[1]))
        self.assertEqual(back['Phone'].iloc[0], '9876543210')   # stays text, not a number
        self.assertEqual(int(back['Excel Row'].iloc[1]), 3)

    def test_strings_are_never_formulas_links_or_numbers(self):
        df = pd.DataFrame({'Name': ['=SUM(A1:A2)', 'http://example.com', '007']})
        back = self.roundtrip({'S': df})['S']
        self.assertEqual(back['Name'].tolist(), ['=SUM(A1:A2)', 'http://example.com', '007'])

    def test_dates_and_missing_numbers(self):
        df = pd.DataFrame({
            'When': [pd.Timestamp('2025-03-01 10:30:00'), pd.NaT],
            'Amount': [1.5, np.nan],
            'Id': [1, 2],            # keeps row 2 from being an entirely empty row
        })
        back = self.roundtrip({'S': df})['S']
        self.assertEqual(pd.Timestamp(back['When'].iloc[0]), pd.Timestamp('2025-03-01 10:30:00'))
        self.assertTrue(pd.isna(back['When'].iloc[1]))
        self.assertEqual(back['Amount'].iloc[0], 1.5)
        self.assertTrue(pd.isna(back['Amount'].iloc[1]))

    def test_several_sheets_and_long_sheet_name(self):
        df = pd.DataFrame({'A': [1]})
        out = self.roundtrip({'First': df, 'Second': df})
        self.assertEqual(set(out), {'First', 'Second'})
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, 'o.xlsx')
            xlsx_writer.write_workbook(path, {'x' * 50: df})
            self.assertEqual(pd.ExcelFile(path, engine='openpyxl').sheet_names, ['x' * 31])


if __name__ == '__main__':
    unittest.main()
