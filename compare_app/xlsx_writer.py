"""
Fast .xlsx writer (xlsxwriter only, no Django imports so it is easy to unit-test).

Why: pandas' DataFrame.to_excel builds a styled cell object for every cell, which is
the slowest step for lakhs of rows. Writing rows straight to xlsxwriter in
constant-memory mode is roughly twice as fast and uses far less RAM.

Safety choices:
  * strings are never turned into formulas / links / numbers ("=SUM(A1)" stays text),
  * missing values (NaN / NaT / None) become empty cells, same as before,
  * the caller falls back to pandas if anything here raises.
"""
import pandas as pd

HEADER_STYLE = {'bold': True, 'bg_color': '#F2F4F7', 'bottom': 1}
DATETIME_FORMAT = 'yyyy-mm-dd hh:mm:ss'
DATE_FORMAT = 'yyyy-mm-dd'


def _clean_values(series):
    """Series -> list of python values with missing values replaced by ''."""
    values = series.astype(object).where(series.notna(), '')
    return values.tolist()


def write_workbook(path, sheets):
    """
    sheets: {sheet_name: DataFrame}. Writes every frame to its own sheet.
    Raises on any problem so the caller can fall back to the slower pandas path.
    """
    import xlsxwriter

    workbook = xlsxwriter.Workbook(str(path), {
        'constant_memory': True,
        'strings_to_formulas': False,
        'strings_to_urls': False,
        'strings_to_numbers': False,
        'nan_inf_to_errors': True,
    })
    try:
        header_fmt = workbook.add_format(HEADER_STYLE)
        datetime_fmt = workbook.add_format({'num_format': DATETIME_FORMAT})
        date_fmt = workbook.add_format({'num_format': DATE_FORMAT})

        for sheet_name, frame in sheets.items():
            ws = workbook.add_worksheet(str(sheet_name)[:31])
            ws.write_row(0, 0, [str(c) for c in frame.columns], header_fmt)

            columns = []
            for position, col in enumerate(frame.columns):
                series = frame.iloc[:, position]
                if pd.api.types.is_datetime64_any_dtype(series):
                    # whole-day values get a date-only format, anything else a full timestamp
                    non_null = series.dropna()
                    midnight_only = bool(len(non_null)) and bool((non_null == non_null.dt.normalize()).all())
                    ws.set_column(position, position, 19 if not midnight_only else 12,
                                  date_fmt if midnight_only else datetime_fmt)
                columns.append(_clean_values(series))

            for row_number, row in enumerate(zip(*columns), start=1):
                ws.write_row(row_number, 0, row)

            if not columns:
                continue
            ws.freeze_panes(1, 0)
    finally:
        workbook.close()
