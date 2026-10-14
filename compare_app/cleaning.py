"""
Pure-pandas cleaning logic for the DataFlow Toolkit.

This module deliberately has NO Django imports, so it can be unit-tested
(and reused later by background jobs) without starting Django.
"""

import pandas as pd

ROW_COL = 'Excel Row'
PHONE_HINTS = ('phone', 'mobile', 'contact')
TITLECASE_HINTS = ('name', 'state', 'city')


# ------------------------------------------------------------------
# Small helpers
# ------------------------------------------------------------------
def find_phone_col(columns):
    """First column whose name looks like a phone column, else None."""
    for col in columns:
        low = col.lower()
        if any(hint in low for hint in PHONE_HINTS):
            return col
    return None


def as_text(series):
    """Series -> plain strings; missing values become ''. Safe for any dtype."""
    missing = series.isna()
    return series.astype(object).where(~missing, '').astype(str)


def _strip_float_suffix(series):
    """Excel stores numbers as floats: '9876543210.0' -> '9876543210'."""
    return as_text(series).str.replace(r'\.0$', '', regex=True)


def _wants_titlecase(col):
    low = col.lower()
    return any(hint in low for hint in TITLECASE_HINTS)


# ------------------------------------------------------------------
# Single file: clean / deduplicate
# ------------------------------------------------------------------
def clean_single(df, columns, remove_duplicates=True, remove_blanks=True,
                 remove_invalid_phones=True, auto_fix=True):
    """
    Clean one DataFrame. NOTE: `df` is modified in place (an 'Excel Row' column is
    inserted) to avoid copying very large files.

    Returns a dict with:
      fresh_df          cleaned rows (includes 'Excel Row')
      fresh_export      cleaned rows without 'Excel Row' (what the user downloads)
      duplicates_df, blanks_df, invalid_phones_df, empty_cells_df   (original rows)
      total_blanks      number of empty cells in the checked columns
      total_rows
    """
    total_rows = len(df)
    df.insert(0, ROW_COL, df.index + 2)  # +2 = header row + 1-based numbering

    phone_col = find_phone_col(columns)
    if phone_col:
        df[phone_col] = _strip_float_suffix(df[phone_col])

    # Display values: trimmed text (+ Title Case for name/state/city when auto_fix)
    df_check = df.copy()
    for col in columns:
        df_check[col] = as_text(df_check[col]).str.strip()
        if auto_fix and _wants_titlecase(col):
            df_check[col] = df_check[col].str.title()

    # One definition of "empty" everywhere: blank or whitespace-only
    empty_mask = df_check[columns].eq('')
    total_blanks = int(empty_mask.sum().sum())
    has_empty_cells = empty_mask.any(axis=1)
    empty_cells_df = df.loc[has_empty_cells]

    valid_indices = df.index
    blanks_df = pd.DataFrame()
    invalid_phones_df = pd.DataFrame()
    duplicates_df = pd.DataFrame()

    # 1. Blank rows
    if remove_blanks:
        blank_indices = df_check[has_empty_cells].index
        blanks_df = df.loc[blank_indices]
        valid_indices = valid_indices.difference(blank_indices)
        df_check = df_check.loc[valid_indices]

    # 2. Phone fixing / invalid phones
    if phone_col:
        only_numbers = df_check[phone_col].str.replace(r'\D', '', regex=True)
        if auto_fix:
            # '91' prefix only when exactly 10 digits follow; leading zeros likewise
            only_numbers = only_numbers.str.replace(r'^91(?=\d{10}$)', '', regex=True)
            only_numbers = only_numbers.str.replace(r'^0+(?=\d{10}$)', '', regex=True)
        df_check[phone_col] = only_numbers

        if remove_invalid_phones:
            is_invalid = (only_numbers.str.len() > 0) & (only_numbers.str.len() != 10)
            invalid_indices = df_check[is_invalid].index
            invalid_phones_df = df.loc[invalid_indices]
            valid_indices = valid_indices.difference(invalid_indices)
            df_check = df_check.loc[valid_indices]

    # 3. Duplicates. Always compared case-insensitively on EVERY checked column
    #    (previously only name/state/city were, when auto_fix was on).
    if remove_duplicates:
        key = df_check[columns].apply(lambda s: s.str.lower())
        is_duplicate = key.duplicated(keep='first')
        duplicate_indices = key[is_duplicate].index
        duplicates_df = df.loc[duplicate_indices]
        valid_indices = valid_indices.difference(duplicate_indices)

    fresh_df = df.loc[valid_indices].copy()
    if auto_fix:
        for col in columns:
            fresh_df[col] = df_check.loc[valid_indices, col]
    fresh_export = fresh_df.drop(columns=[ROW_COL])

    return {
        'total_rows': total_rows,
        'fresh_df': fresh_df,
        'fresh_export': fresh_export,
        'duplicates_df': duplicates_df,
        'blanks_df': blanks_df,
        'invalid_phones_df': invalid_phones_df,
        'empty_cells_df': empty_cells_df,
        'total_blanks': total_blanks,
    }


# ------------------------------------------------------------------
# Two files: find matching rows
# ------------------------------------------------------------------
def compare_frames(df1, df2, columns):
    """
    Exact matches between two DataFrames on `columns` (trimmed, case-insensitive).
    Both frames get a 'File N Row' column inserted in place.
    Returns (matches_df, total_blank_cells).
    """
    df1.insert(0, 'File 1 Row', df1.index + 2)
    df2.insert(0, 'File 2 Row', df2.index + 2)

    phone_col = find_phone_col(columns)
    if phone_col:
        for d in (df1, df2):
            d[phone_col] = _strip_float_suffix(d[phone_col])

    cleaned = []
    for d in (df1, df2):
        c = d.copy()
        for col in columns:
            c[col] = as_text(c[col]).str.strip().str.lower()
        cleaned.append(c)
    df1_clean, df2_clean = cleaned

    total_blanks = int(
        df1_clean[columns].eq('').sum().sum() + df2_clean[columns].eq('').sum().sum()
    )
    matches = pd.merge(df1_clean, df2_clean, on=columns, how='inner')
    return matches, total_blanks
