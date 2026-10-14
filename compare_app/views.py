import logging
import math
import re
import time
import uuid
from pathlib import Path

import pandas as pd
from django.conf import settings
from django.contrib.auth import login, logout
from django.contrib.auth.forms import AuthenticationForm
from django.db.models import F
from django.http import FileResponse, Http404
from django.shortcuts import redirect, render
from django.urls import reverse

from . import cleaning
from .forms import RegisterForm
from .models import UserProfile

logger = logging.getLogger(__name__)

# ==========================================
# CONFIG
# ==========================================
FREE_ROW_LIMIT = 1000
PRO_ROW_LIMIT = 1048575   # Excel max data rows (1,048,576 - 1 header)
FREE_MAX_FILE_MB = 5
PRO_MAX_FILE_MB = 150
ALLOWED_EXTENSIONS = ('.xlsx', '.xls', '.csv')
RESULT_MAX_AGE_SECONDS = 24 * 60 * 60  # results 24 ghante baad delete

RESULT_DIR = Path(settings.BASE_DIR) / 'processed_files'
RESULT_DIR.mkdir(exist_ok=True)

TOKEN_RE = re.compile(r'^[0-9a-f]{32}$')

# Contact details come from settings.py / environment variables (SUPPORT_EMAIL, WHATSAPP_NUMBER)
SUPPORT_EMAIL = getattr(settings, 'SUPPORT_EMAIL', '')
WHATSAPP_NUMBER = getattr(settings, 'WHATSAPP_NUMBER', '')


class LimitReached(ValueError):
    """Free plan ki limit cross hone par (upgrade prompt dikhane ke liye)."""
    pass


# ==========================================
# HELPERS
# ==========================================
def get_user_plan(user):
    """Returns (is_pro, profile). Profile na ho toh khud bana deta hai."""
    if not user.is_authenticated:
        return False, None
    profile, _ = UserProfile.objects.get_or_create(user=user)
    return profile.has_active_pro(), profile


def validate_upload(file_obj, is_pro):
    name = file_obj.name.lower()
    if not name.endswith(ALLOWED_EXTENSIONS):
        raise ValueError("Only .xlsx, .xls or .csv files are allowed.")
    max_mb = PRO_MAX_FILE_MB if is_pro else FREE_MAX_FILE_MB
    if file_obj.size > max_mb * 1024 * 1024:
        msg = f"File too large. {'Pro' if is_pro else 'Free'} plan supports files up to {max_mb} MB."
        raise (ValueError if is_pro else LimitReached)(msg)


def check_row_limit(total_rows, is_pro, request):
    limit = PRO_ROW_LIMIT if is_pro else FREE_ROW_LIMIT
    if total_rows > limit:
        if is_pro:
            raise ValueError(f"File has {total_rows:,} rows. Maximum {limit:,} rows allowed.")
        raise LimitReached(
            f"Your file has {total_rows:,} rows, but the Free plan supports up to {limit:,} rows per file."
        )


def parse_columns(columns_input):
    cols = [c.strip() for c in (columns_input or '').split(',') if c.strip()]
    if not cols:
        raise ValueError("Please enter at least one column name.")
    # duplicate names hatao, order same rakho
    return list(dict.fromkeys(cols))


def read_table(file_obj, columns=None):
    """Excel (.xlsx/.xls) ya CSV padhta hai. columns diye to sirf wahi columns lauta hai."""
    name = file_obj.name.lower()

    if not name.endswith('.csv'):
        if columns:
            return pd.read_excel(file_obj, engine='calamine', usecols=columns)
        return pd.read_excel(file_obj, engine='calamine')

    # ---- CSV ----
    head = file_obj.read(4096)
    file_obj.seek(0)
    first_line = head.decode('utf-8', errors='ignore').splitlines()[0] if head else ''
    sep = max([',', ';', '\t', '|'], key=first_line.count)

    df = None
    for enc in ('utf-8-sig', 'cp1252', 'latin-1'):
        try:
            file_obj.seek(0)
            df = pd.read_csv(file_obj, dtype=str, sep=sep, encoding=enc)
            break
        except UnicodeDecodeError:
            continue
        except pd.errors.EmptyDataError:
            raise ValueError("The uploaded CSV file is empty.")
        except pd.errors.ParserError:
            raise ValueError("Could not read this CSV. Please check that it is a valid CSV file.")
    if df is None:
        raise ValueError("Could not read this CSV file (unsupported encoding).")

    df.columns = df.columns.astype(str).str.strip()
    if columns:
        missing = [c for c in columns if c not in df.columns]
        if missing:
            raise ValueError(f"Columns not found: {missing}. Available columns: {list(df.columns)}")
        df = df[columns]
    return df


def cleanup_old_files():
    now = time.time()
    for f in RESULT_DIR.glob('*.xlsx'):
        try:
            if now - f.stat().st_mtime > RESULT_MAX_AGE_SECONDS:
                f.unlink()
        except OSError:
            pass


def excel_engine():
    """xlsxwriter is several times faster than openpyxl for big files; fall back if missing."""
    try:
        import xlsxwriter  # noqa: F401
        return 'xlsxwriter'
    except ImportError:
        return 'openpyxl'


def save_excel(request, sheets, download_name, allow_empty=False):
    """
    sheets: {sheet_name: DataFrame}. File save karke download URL return karta hai.
    Token session mein store hota hai, sirf same user download kar sakta hai.
    """
    if not allow_empty:
        sheets = {n: d for n, d in sheets.items() if d is not None and not d.empty}
    if not sheets:
        return None

    token = uuid.uuid4().hex
    path = RESULT_DIR / f"{token}.xlsx"
    with pd.ExcelWriter(path, engine=excel_engine()) as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    files = request.session.get('result_files', {})
    files[token] = download_name
    # sirf last 50 files yaad rakho
    if len(files) > 50:
        files = dict(list(files.items())[-50:])
    request.session['result_files'] = files
    return reverse('download', args=[token])


def to_html(df, css_class):
    if df is None or df.empty:
        return None
    return df.head(100).to_html(
        classes=f'table {css_class} table-striped mb-0',
        index=False, justify='left', na_rep=''
    )


def add_usage(user, rows):
    if user.is_authenticated and rows > 0:
        UserProfile.objects.get_or_create(user=user)
        UserProfile.objects.filter(user=user).update(
            total_rows_processed=F('total_rows_processed') + rows
        )


def format_duration(seconds):
    mins, secs = int(seconds // 60), round(seconds % 60, 2)
    return f"{mins} min {secs} sec"


# ==========================================
# 1. LANDING PAGE
# ==========================================
def landing_page(request):
    if request.user.is_authenticated:
        return redirect('app')
    return render(request, 'compare_app/landing.html', {'support_email': SUPPORT_EMAIL})


# ==========================================
# 2. REGISTER
# ==========================================
def register(request):
    if request.user.is_authenticated:
        return redirect('app')
    if request.method == 'POST':
        form = RegisterForm(request.POST)
        if form.is_valid():
            user = form.save()
            login(request, user)
            return redirect('app')
    else:
        form = RegisterForm()
    return render(request, 'compare_app/register.html', {'form': form})


# ==========================================
# 3. LOGIN
# ==========================================
def user_login(request):
    if request.user.is_authenticated:
        return redirect('app')
    if request.method == 'POST':
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            login(request, form.get_user())
            return redirect('app')
    else:
        form = AuthenticationForm()
    return render(request, 'compare_app/login.html', {'form': form})


# ==========================================
# 4. LOGOUT
# ==========================================
def user_logout(request):
    logout(request)
    return redirect('home')


# ==========================================
# 5. SECURE DOWNLOAD
# ==========================================
def download_result(request, token):
    if not TOKEN_RE.match(token):
        raise Http404
    files = request.session.get('result_files', {})
    if token not in files:
        raise Http404
    path = RESULT_DIR / f"{token}.xlsx"
    if not path.exists():
        raise Http404("File expired. Please process again.")
    return FileResponse(open(path, 'rb'), as_attachment=True, filename=files[token])


# ==========================================
# 6. MAIN DASHBOARD
# ==========================================
def excel_dashboard(request):
    is_pro, profile = get_user_plan(request.user)
    context = {
        'active_tab': 'single',
        'is_pro': is_pro,
        'free_row_limit': FREE_ROW_LIMIT,
        'pro_max_mb': PRO_MAX_FILE_MB,
        'total_rows_processed': profile.total_rows_processed if profile else 0,
    }

    if request.method == 'POST':
        mode = request.POST.get('mode')
        if mode not in ('single', 'compare'):
            mode = 'single'
        context['active_tab'] = mode

        try:
            cleanup_old_files()
            columns_to_check = parse_columns(request.POST.get('columns', 'Name,Phone,State'))

            # ==========================================
            # MODE 1: SINGLE FILE
            # ==========================================
            if mode == 'single':
                file_single = request.FILES.get('file_single')
                if not file_single:
                    raise ValueError("Please upload an Excel file.")
                validate_upload(file_single, is_pro)

                remove_duplicates = request.POST.get('remove_duplicates') == 'yes'
                remove_blanks = request.POST.get('remove_blanks') == 'yes'
                remove_invalid_phones = request.POST.get('remove_invalid_phones') == 'yes'
                auto_fix = request.POST.get('auto_fix') == 'yes'

                start_time = time.perf_counter()

                df = read_table(file_single)
                df.columns = df.columns.astype(str).str.strip()
                df = df.dropna(how='all')   # fully empty (formatted-only) rows are not data
                total_rows = len(df)

                if total_rows == 0:
                    raise ValueError("The uploaded file has no data rows.")
                check_row_limit(total_rows, is_pro, request)

                missing_cols = [c for c in columns_to_check if c not in df.columns]
                if missing_cols:
                    raise ValueError(
                        f"Columns not found: {missing_cols}. "
                        f"Available columns: {list(df.columns)}"
                    )

                result = cleaning.clean_single(
                    df, columns_to_check,
                    remove_duplicates=remove_duplicates,
                    remove_blanks=remove_blanks,
                    remove_invalid_phones=remove_invalid_phones,
                    auto_fix=auto_fix,
                )
                fresh_df = result['fresh_df']
                fresh_df_export = result['fresh_export']
                duplicates_df = result['duplicates_df']
                blanks_df = result['blanks_df']
                invalid_phones_df = result['invalid_phones_df']
                empty_cells_df = result['empty_cells_df']
                total_blanks = result['total_blanks']

                health_score = math.floor((len(fresh_df) / total_rows) * 100) if total_rows > 0 else 0
                if health_score >= 90:
                    health_color = 'success'
                elif health_score >= 70:
                    health_color = 'warning'
                else:
                    health_color = 'danger'

                main_sheets = {'Fresh Clean Data': fresh_df_export}
                if not duplicates_df.empty:
                    main_sheets['Found Duplicates'] = duplicates_df
                main_url = save_excel(request, main_sheets, 'Fresh_Clean_Data.xlsx', allow_empty=True)

                # Success ke baad hi usage count karo
                add_usage(request.user, total_rows)

                context.update({
                    'single_success': True,
                    's_health_score': health_score,
                    's_health_color': health_color,
                    's_total_rows': total_rows,
                    's_duplicate_count': len(duplicates_df),
                    's_blank_rows_removed': len(blanks_df),
                    's_invalid_phones': len(invalid_phones_df),
                    's_total_blanks': total_blanks,
                    's_fresh_count': len(fresh_df),
                    's_download_uri': main_url,
                    's_execution_time': format_duration(time.perf_counter() - start_time),

                    's_duplicates_table': to_html(duplicates_df, 'table-warning'),
                    's_blanks_table': to_html(blanks_df, 'table-danger'),
                    's_invalid_table': to_html(invalid_phones_df, 'table-info'),
                    's_empty_cells_table': to_html(empty_cells_df, 'table-secondary'),

                    's_dup_uri': save_excel(request, {'Removed Duplicates': duplicates_df}, 'Removed_Duplicates.xlsx'),
                    's_blank_uri': save_excel(request, {'Blank Rows': blanks_df}, 'Removed_Blank_Rows.xlsx'),
                    's_invalid_uri': save_excel(request, {'Invalid Phones': invalid_phones_df}, 'Removed_Invalid_Phones.xlsx'),
                    's_empty_uri': save_excel(request, {'Rows with Empty Cells': empty_cells_df}, 'Rows_With_Missing_Data.xlsx'),
                })

            # ==========================================
            # MODE 2: COMPARE
            # ==========================================
            elif mode == 'compare':
                file1 = request.FILES.get('file1')
                file2 = request.FILES.get('file2')
                if not file1 or not file2:
                    raise ValueError("Please upload both files.")
                validate_upload(file1, is_pro)
                validate_upload(file2, is_pro)

                start_time = time.perf_counter()

                try:
                    df1 = read_table(file1, columns_to_check)
                    df2 = read_table(file2, columns_to_check)
                except ValueError as e:
                    msg = str(e)
                    if 'Columns not found' in msg or 'CSV' in msg or 'encoding' in msg:
                        raise
                    raise ValueError(
                        f"One or both files do not contain all these columns: {columns_to_check}"
                    )
                df1 = df1.dropna(how='all')   # fully empty (formatted-only) rows are not data
                df2 = df2.dropna(how='all')
                check_row_limit(max(len(df1), len(df2)), is_pro, request)

                if is_pro and (len(df1) + len(df2)) > 2000000:
                    raise ValueError(
                        f"Combined rows ({len(df1) + len(df2):,}) exceed the compare limit of 20,00,000 rows."
                    )
                if len(df1) == 0 or len(df2) == 0:
                    raise ValueError("One of the uploaded files has no data rows.")

                duplicates, total_blanks = cleaning.compare_frames(df1, df2, columns_to_check)
                if len(duplicates) > PRO_ROW_LIMIT:
                    raise ValueError(
                        f"The two files have {len(duplicates):,} matching rows, which is more than one Excel "
                        f"sheet can hold. Try matching on more columns so each row is more specific."
                    )

                if duplicates.empty:
                    main_sheets = {'Result': pd.DataFrame({'Message': ['No duplicates found']})}
                else:
                    main_sheets = {'Exact Duplicates': duplicates}
                main_url = save_excel(request, main_sheets, 'Comparison_Full_Report.xlsx', allow_empty=True)

                add_usage(request.user, len(df1) + len(df2))

                context.update({
                    'compare_success': True,
                    'c_total_rows_1': len(df1),
                    'c_total_rows_2': len(df2),
                    'c_duplicate_count': len(duplicates),
                    'c_total_blanks': total_blanks,
                    'c_download_uri': main_url,
                    'c_execution_time': format_duration(time.perf_counter() - start_time),
                    'c_duplicates_table': to_html(duplicates, 'table-success'),
                    'c_dup_uri': save_excel(request, {'Exact Duplicates': duplicates}, 'Matched_Exact_Duplicates.xlsx'),
                })

            # usage number refresh (dashboard par dikhane ke liye)
            if profile:
                profile.refresh_from_db()
                context['total_rows_processed'] = profile.total_rows_processed

        except LimitReached as e:
            # Free plan limit: the template shows an Upgrade card instead of a plain error
            context['limit_error'] = str(e)
        except ValueError as e:
            context['error'] = str(e)
        except Exception:
            logger.exception("Unexpected error while processing file")
            context['error'] = (
                "Something went wrong while processing your file. "
                "Please check the file format and try again."
            )

    return render(request, 'compare_app/compare.html', context)

def upgrade(request):
    is_pro, profile = get_user_plan(request.user)
    return render(request, 'compare_app/upgrade.html', {
        'is_pro': is_pro,
        'free_row_limit': FREE_ROW_LIMIT,
        'pro_row_limit': PRO_ROW_LIMIT,
        'free_max_mb': FREE_MAX_FILE_MB,
        'pro_max_mb': PRO_MAX_FILE_MB,
        'support_email': SUPPORT_EMAIL,
        'whatsapp_number': WHATSAPP_NUMBER,
    })