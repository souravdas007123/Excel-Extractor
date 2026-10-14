"""Unit tests for compare_app.cleaning (pure pandas, no database needed)."""
import unittest

import numpy as np
import pandas as pd

from compare_app.cleaning import as_text, clean_single, compare_frames, find_phone_col

COLS = ['Name', 'Phone', 'State']


def make_df(rows):
    return pd.DataFrame(rows, columns=COLS)


class HelperTests(unittest.TestCase):
    def test_find_phone_col(self):
        self.assertEqual(find_phone_col(['Name', 'Mobile No', 'State']), 'Mobile No')
        self.assertIsNone(find_phone_col(['Name', 'State']))

    def test_as_text_handles_nan_float_and_int(self):
        out = as_text(pd.Series([9876543210.0, np.nan, 5.0])).tolist()
        self.assertEqual(out, ['9876543210.0', '', '5.0'])
        self.assertEqual(as_text(pd.Series([1, 2])).tolist(), ['1', '2'])


class CleanSingleTests(unittest.TestCase):
    def test_phone_prefixes_and_title_case(self):
        df = make_df([
            ['rahul sharma', '+91 98765 43210', 'delhi'],
            ['neha gupta', '09123456780', 'bihar'],
        ])
        r = clean_single(df, COLS)
        self.assertEqual(r['fresh_export']['Phone'].tolist(), ['9876543210', '9123456780'])
        self.assertEqual(r['fresh_export']['Name'].tolist(), ['Rahul Sharma', 'Neha Gupta'])
        self.assertEqual(r['fresh_export']['State'].tolist(), ['Delhi', 'Bihar'])

    def test_float_phone_from_excel(self):
        df = make_df([['A B', 9876543210.0, 'UP'], ['C D', np.nan, 'UP']])
        r = clean_single(df, COLS, remove_blanks=False)
        self.assertEqual(r['fresh_export']['Phone'].tolist(), ['9876543210', ''])

    def test_duplicates_removed_case_insensitively(self):
        df = make_df([
            ['rahul sharma', '+91 98765 43210', 'delhi'],
            ['Rahul Sharma', '9876543210', 'Delhi'],
        ])
        r = clean_single(df, COLS)
        self.assertEqual(len(r['fresh_df']), 1)
        self.assertEqual(len(r['duplicates_df']), 1)
        self.assertEqual(int(r['duplicates_df']['Excel Row'].iloc[0]), 3)  # 2nd data row

    def test_duplicates_in_non_name_columns_ignore_case(self):
        # Regression: with auto_fix on, a column like Email used to stay case-sensitive.
        df = pd.DataFrame(
            [['A', 'A@X.COM'], ['A', 'a@x.com']], columns=['Name', 'Email']
        )
        r = clean_single(df, ['Name', 'Email'], auto_fix=True)
        self.assertEqual(len(r['fresh_df']), 1)
        self.assertEqual(len(r['duplicates_df']), 1)

    def test_blank_and_invalid_phone_rows_are_separated(self):
        df = make_df([
            ['Priya Singh', '98765', 'Punjab'],      # invalid phone
            ['Amit Verma', '', 'UP'],                # blank phone
            ['Good One', '9999999999', 'Goa'],
        ])
        r = clean_single(df, COLS)
        self.assertEqual(len(r['invalid_phones_df']), 1)
        self.assertEqual(len(r['blanks_df']), 1)
        self.assertEqual(r['fresh_export']['Name'].tolist(), ['Good One'])

    def test_missing_cell_count_matches_flagged_rows(self):
        # Regression: whitespace-only cells are blank for BOTH the count and the rows.
        df = make_df([
            ['   ', '9999999999', 'Goa'],
            ['Real', '8888888888', 'UP'],
            [None, None, 'UP'],
        ])
        r = clean_single(df, COLS)
        self.assertEqual(r['total_blanks'], 3)
        self.assertEqual(len(r['empty_cells_df']), 2)
        self.assertEqual(len(r['blanks_df']), 2)

    def test_no_autofix_keeps_original_values(self):
        df = make_df([['rahul sharma', '+91 98765 43210', 'delhi']])
        r = clean_single(df, COLS, auto_fix=False, remove_invalid_phones=False)
        self.assertEqual(r['fresh_export']['Name'].tolist(), ['rahul sharma'])
        self.assertEqual(r['fresh_export']['Phone'].tolist(), ['+91 98765 43210'])

    def test_excel_row_numbers_survive_empty_row_drop(self):
        df = make_df([
            ['A One', '9111111111', 'UP'],
            [None, None, None],                      # fully empty -> dropped by caller
            ['B Two', '9222222222', 'UP'],
        ]).dropna(how='all')
        r = clean_single(df, COLS)
        self.assertEqual(r['fresh_df']['Excel Row'].tolist(), [2, 4])

    def test_extra_columns_are_kept(self):
        df = pd.DataFrame(
            [['A One', '9111111111', 'UP', 'KEEP ME']], columns=COLS + ['Notes']
        )
        r = clean_single(df, COLS)
        self.assertEqual(r['fresh_export']['Notes'].tolist(), ['KEEP ME'])


class CompareTests(unittest.TestCase):
    def test_matches_are_case_and_space_insensitive(self):
        a = make_df([['Rahul Sharma', '9876543210', 'Delhi'], ['X Y', '1111111111', 'UP']])
        b = make_df([['  rahul sharma ', 9876543210.0, 'DELHI']])
        matches, blanks = compare_frames(a, b, COLS)
        self.assertEqual(len(matches), 1)
        self.assertEqual(blanks, 0)

    def test_blank_count_uses_trimmed_emptiness(self):
        a = make_df([['  ', '1', 'UP']])
        b = make_df([[None, '2', 'UP']])
        _, blanks = compare_frames(a, b, COLS)
        self.assertEqual(blanks, 2)

    def test_no_matches(self):
        a = make_df([['A', '1', 'X']])
        b = make_df([['B', '2', 'Y']])
        matches, _ = compare_frames(a, b, COLS)
        self.assertTrue(matches.empty)


if __name__ == '__main__':
    unittest.main()
