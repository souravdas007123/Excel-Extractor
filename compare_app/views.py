import json
import logging
import math
import os
import re
import threading
import time
import uuid
from pathlib import Path

import pandas as pd
from django.conf import settings
from django.contrib.auth import login, logout
from django.contrib.auth.forms import AuthenticationForm
from django.db.models import F
from django.http import FileResponse, Http404, JsonResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.views.decorators.http import require_GET, require_POST

from . import cleaning, xlsx_writer
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
JOB_DIR = RESULT_DIR / 'jobs'            # one small .json status file per background job
INCOMING_DIR = RESULT_DIR / 'incoming'   # uploaded files waiting for / being processed by a job
JOB_DIR.mkdir(exist_ok=True)
INCOMING_DIR.mkdir(exist_ok=True)

JOB_STALE_SECONDS = 20 * 60   # a "running" job that has not reported progress for this long is treated as failed
JOB_SLOTS = threading.BoundedSemaphore(max(1, int(getattr(settings, 'MAX_CONCURRENT_JOBS', 2))))
_JOB_LOCK = threading.Lock()

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
    """Delete results, job status files and leftover uploads older than 24 hours."""
    now = time.time()
    for pattern_dir, pattern in ((RESULT_DIR, '*.xlsx'), (JOB_DIR, '*.json'), (INCOMING_DIR, '*')):
        for f in pattern_dir.glob(pattern):
            try:
                if f.is_file() and now - f.stat().st_mtime > RESULT_MAX_AGE_SECONDS:
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


def write_excel_file(path, sheets):
    """
    Fast path: xlsx_writer (about 2x faster, low memory). If it fails for any reason
    (odd data types etc.) we delete the partial file and use the slower pandas writer.
    """
    if excel_engine() == 'xlsxwriter':
        try:
            xlsx_writer.write_workbook(path, sheets)
            return
        except Exception:
            logger.exception("Fast Excel writer failed, falling back to pandas")
            try:
                Path(path).unlink()
            except OSError:
                pass
        kwargs = {'engine_kwargs': {'options': {'strings_to_formulas': False, 'strings_to_urls': False}}}
    else:
        kwargs = {}
    with pd.ExcelWriter(path, engine=excel_engine(), **kwargs) as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)


def store_excel(sheets, allow_empty=False):
    """Write sheets to a new result file and return its token (None if there is nothing to write)."""
    if not allow_empty:
        sheets = {n: d for n, d in sheets.items() if d is not None and not d.empty}
    if not sheets:
        return None
    token = uuid.uuid4().hex
    write_excel_file(RESULT_DIR / f"{token}.xlsx", sheets)
    return token


def remember_files(request, new_files):
    """Allow this browser session to download the given {token: download_name} files."""
    files = request.session.get('result_files', {})
    files.update(new_files)
    if len(files) > 50:                      # only keep the last 50
        files = dict(list(files.items())[-50:])
    request.session['result_files'] = files


def save_excel(request, sheets, download_name, allow_empty=False):
    """
    sheets: {sheet_name: DataFrame}. Saves the file and returns its download URL.
    The token is stored in the session, so only the same user can download it.
    """
    token = store_excel(sheets, allow_empty)
    if token is None:
        return None
    remember_files(request, {token: download_name})
    return reverse('download', args=[token])


def to_html(df, css_class):
    if df is None or df.empty:
        return None
    return df.head(100).to_html(
        classes=f'table {css_class} table-striped mb-0',
        index=False, justify='left', na_rep=''
    )


def add_usage_for_id(user_id, rows):
    if user_id and rows > 0:
        UserProfile.objects.get_or_create(user_id=user_id)
        UserProfile.objects.filter(user_id=user_id).update(
            total_rows_processed=F('total_rows_processed') + rows
        )


def add_usage(user, rows):
    if user.is_authenticated:
        add_usage_for_id(user.pk, rows)


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
# 6. PROCESSING (shared by the normal page and the background jobs)
# ==========================================
def _no_progress(percent, message):
    pass


def process_single(file_obj, post, is_pro, saver, progress=_no_progress):
    """
    Clean one file. `saver(sheets, download_name, allow_empty=False)` stores an Excel file
    and returns its download URL (or None). Returns (context_update, rows_processed).
    Raises ValueError / LimitReached for problems the user should see.
    """
    columns_to_check = parse_columns(post.get('columns', 'Name,Phone,State'))
    remove_duplicates = post.get('remove_duplicates') == 'yes'
    remove_blanks = post.get('remove_blanks') == 'yes'
    remove_invalid_phones = post.get('remove_invalid_phones') == 'yes'
    auto_fix = post.get('auto_fix') == 'yes'

    start_time = time.perf_counter()

    progress(8, "Reading your file...")
    df = read_table(file_obj)
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how='all')   # fully empty (formatted-only) rows are not data
    total_rows = len(df)

    if total_rows == 0:
        raise ValueError("The uploaded file has no data rows.")
    check_row_limit(total_rows, is_pro, None)

    missing_cols = [c for c in columns_to_check if c not in df.columns]
    if missing_cols:
        raise ValueError(
            f"Columns not found: {missing_cols}. "
            f"Available columns: {list(df.columns)}"
        )

    progress(30, f"Cleaning {total_rows:,} rows...")
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

    progress(55, "Creating the clean Excel file...")
    main_sheets = {'Fresh Clean Data': fresh_df_export}
    if not duplicates_df.empty:
        main_sheets['Found Duplicates'] = duplicates_df
    main_url = saver(main_sheets, 'Fresh_Clean_Data.xlsx', allow_empty=True)

    progress(75, "Creating report files...")
    update = {
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

        's_duplicates_table': to_html(duplicates_df, 'table-warning'),
        's_blanks_table': to_html(blanks_df, 'table-danger'),
        's_invalid_table': to_html(invalid_phones_df, 'table-info'),
        's_empty_cells_table': to_html(empty_cells_df, 'table-secondary'),

        's_dup_uri': saver({'Removed Duplicates': duplicates_df}, 'Removed_Duplicates.xlsx'),
        's_blank_uri': saver({'Blank Rows': blanks_df}, 'Removed_Blank_Rows.xlsx'),
        's_invalid_uri': saver({'Invalid Phones': invalid_phones_df}, 'Removed_Invalid_Phones.xlsx'),
        's_empty_uri': saver({'Rows with Empty Cells': empty_cells_df}, 'Rows_With_Missing_Data.xlsx'),
    }
    update['s_execution_time'] = format_duration(time.perf_counter() - start_time)
    return update, total_rows


def process_compare(file1, file2, post, is_pro, saver, progress=_no_progress):
    """Find rows that exist in both files. Same contract as process_single."""
    columns_to_check = parse_columns(post.get('columns', 'Name,Phone,State'))
    start_time = time.perf_counter()

    progress(8, "Reading both files...")
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
    check_row_limit(max(len(df1), len(df2)), is_pro, None)

    if is_pro and (len(df1) + len(df2)) > 2000000:
        raise ValueError(
            f"Combined rows ({len(df1) + len(df2):,}) exceed the compare limit of 20,00,000 rows."
        )
    if len(df1) == 0 or len(df2) == 0:
        raise ValueError("One of the uploaded files has no data rows.")

    progress(35, f"Comparing {len(df1):,} and {len(df2):,} rows...")
    duplicates, total_blanks = cleaning.compare_frames(df1, df2, columns_to_check)
    if len(duplicates) > PRO_ROW_LIMIT:
        raise ValueError(
            f"The two files have {len(duplicates):,} matching rows, which is more than one Excel "
            f"sheet can hold. Try matching on more columns so each row is more specific."
        )

    progress(60, "Creating the Excel report...")
    if duplicates.empty:
        main_sheets = {'Result': pd.DataFrame({'Message': ['No duplicates found']})}
    else:
        main_sheets = {'Exact Duplicates': duplicates}
    main_url = saver(main_sheets, 'Comparison_Full_Report.xlsx', allow_empty=True)

    progress(85, "Creating the matched rows file...")
    update = {
        'compare_success': True,
        'c_total_rows_1': len(df1),
        'c_total_rows_2': len(df2),
        'c_duplicate_count': len(duplicates),
        'c_total_blanks': total_blanks,
        'c_download_uri': main_url,
        'c_duplicates_table': to_html(duplicates, 'table-success'),
        'c_dup_uri': saver({'Exact Duplicates': duplicates}, 'Matched_Exact_Duplicates.xlsx'),
    }
    update['c_execution_time'] = format_duration(time.perf_counter() - start_time)
    return update, len(df1) + len(df2)


def base_context(request):
    is_pro, profile = get_user_plan(request.user)
    return {
        'active_tab': 'single',
        'is_pro': is_pro,
        'free_row_limit': FREE_ROW_LIMIT,
        'pro_max_mb': PRO_MAX_FILE_MB,
        'total_rows_processed': profile.total_rows_processed if profile else 0,
    }


# ==========================================
# 7. MAIN DASHBOARD (normal form post = works without JavaScript too)
# ==========================================
def excel_dashboard(request):
    is_pro, profile = get_user_plan(request.user)
    context = base_context(request)

    if request.method == 'POST':
        mode = request.POST.get('mode')
        if mode not in ('single', 'compare'):
            mode = 'single'
        context['active_tab'] = mode

        def saver(sheets, name, allow_empty=False):
            return save_excel(request, sheets, name, allow_empty)

        try:
            cleanup_old_files()
            parse_columns(request.POST.get('columns', 'Name,Phone,State'))

            if mode == 'single':
                file_single = request.FILES.get('file_single')
                if not file_single:
                    raise ValueError("Please upload an Excel file.")
                validate_upload(file_single, is_pro)
                update, rows = process_single(file_single, request.POST, is_pro, saver)
            else:
                file1 = request.FILES.get('file1')
                file2 = request.FILES.get('file2')
                if not file1 or not file2:
                    raise ValueError("Please upload both files.")
                validate_upload(file1, is_pro)
                validate_upload(file2, is_pro)
                update, rows = process_compare(file1, file2, request.POST, is_pro, saver)

            add_usage(request.user, rows)   # only counted after a successful run
            context.update(update)

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


# ==========================================
# 8. BACKGROUND JOBS
#    The page uploads the file with progress, gets a job id back immediately, then polls
#    job_status. Work runs in a background thread; status lives in a small JSON file so any
#    gunicorn worker can answer the polling requests. No Redis / Celery needed.
# ==========================================
def _job_path(job_id):
    return JOB_DIR / f"{job_id}.json"


def write_job(job_id, data):
    path = _job_path(job_id)
    tmp = path.with_suffix('.tmp')
    with _JOB_LOCK:
        tmp.write_text(json.dumps(data), encoding='utf-8')
        for attempt in range(8):            # Windows can briefly lock a file that is being read
            try:
                os.replace(tmp, path)
                return
            except PermissionError:
                time.sleep(0.03 * (attempt + 1))
        os.replace(tmp, path)


def load_job(job_id):
    if not TOKEN_RE.match(job_id or ''):
        return None
    path = _job_path(job_id)
    data = None
    for attempt in range(5):
        try:
            data = json.loads(path.read_text(encoding='utf-8'))
            break
        except FileNotFoundError:
            return None
        except (OSError, ValueError):
            time.sleep(0.03 * (attempt + 1))
    if data is None:
        return None
    if data.get('status') in ('queued', 'running') and time.time() - data.get('updated', 0) > JOB_STALE_SECONDS:
        data.update(status='error', progress=100,
                    message="Processing was interrupted (the server may have restarted). Please upload the file again.")
    return data


def update_job(job_id, **changes):
    data = load_job(job_id) or {}
    data.update(changes)
    data['updated'] = time.time()
    write_job(job_id, data)


def job_owner(request):
    if request.user.is_authenticated:
        return f"u{request.user.pk}"
    if not request.session.session_key:
        request.session.create()
    return f"s{request.session.session_key}"


def get_owned_job(request, job_id):
    job = load_job(job_id)
    if job and job.get('owner') == job_owner(request):
        return job
    return None


def run_job(job_id, mode, paths, post, is_pro, user_id, threaded=True):
    """Runs in a background thread (or inline in tests). Never raises; failures go into the job file."""
    handles = []
    try:
        update_job(job_id, status='queued', progress=2, message="Waiting for a free processor...")
        with JOB_SLOTS:
            update_job(job_id, status='running', progress=4, message="Starting...")
            files = {}

            def saver(sheets, name, allow_empty=False):
                token = store_excel(sheets, allow_empty)
                if token is None:
                    return None
                files[token] = name
                return reverse('download', args=[token])

            def progress(percent, message):
                update_job(job_id, progress=percent, message=message)

            handles = [open(p, 'rb') for p in paths]
            if mode == 'single':
                update, rows = process_single(handles[0], post, is_pro, saver, progress)
            else:
                update, rows = process_compare(handles[0], handles[1], post, is_pro, saver, progress)

            add_usage_for_id(user_id, rows)   # only counted after a successful run
            update_job(job_id, status='done', progress=100, message="Done", context=update, files=files)
    except LimitReached as e:
        update_job(job_id, status='limit', progress=100, message=str(e))
    except ValueError as e:
        update_job(job_id, status='error', progress=100, message=str(e))
    except Exception:
        logger.exception("Unexpected error in background job %s", job_id)
        update_job(job_id, status='error', progress=100,
                   message="Something went wrong while processing your file. "
                           "Please check the file format and try again.")
    finally:
        for h in handles:
            try:
                h.close()
            except OSError:
                pass
        for p in paths:
            try:
                os.remove(p)
            except OSError:
                pass
        if threaded:
            from django.db import connection
            connection.close()


def _job_response(job_id):
    return JsonResponse({
        'ok': True,
        'job_id': job_id,
        'status_url': reverse('job_status', args=[job_id]),
        'result_url': reverse('job_result', args=[job_id]),
    })


@require_POST
def start_job(request):
    """Validate the upload, save it, start the job. Always answers with a job id (even for validation errors)."""
    is_pro, _profile = get_user_plan(request.user)
    owner = job_owner(request)
    job_id = uuid.uuid4().hex
    mode = request.POST.get('mode')
    if mode not in ('single', 'compare'):
        mode = 'single'
    now = time.time()
    job = {'id': job_id, 'owner': owner, 'mode': mode, 'status': 'queued', 'progress': 1,
           'message': "Upload received", 'created': now, 'updated': now, 'context': {}, 'files': {}}

    paths = []
    try:
        active = load_job(request.session.get('active_job'))
        if active and active.get('owner') == owner and active.get('status') in ('queued', 'running'):
            raise ValueError("Your previous file is still being processed. Please wait for it to finish.")

        cleanup_old_files()
        parse_columns(request.POST.get('columns', 'Name,Phone,State'))

        if mode == 'single':
            uploads = [request.FILES.get('file_single')]
            if not uploads[0]:
                raise ValueError("Please upload an Excel file.")
        else:
            uploads = [request.FILES.get('file1'), request.FILES.get('file2')]
            if not all(uploads):
                raise ValueError("Please upload both files.")
        for f in uploads:
            validate_upload(f, is_pro)

        for index, f in enumerate(uploads, start=1):
            ext = Path(f.name.lower()).suffix
            path = INCOMING_DIR / f"{job_id}_{index}{ext}"
            with open(path, 'wb') as out:
                for chunk in f.chunks():
                    out.write(chunk)
            paths.append(path)
    except LimitReached as e:
        for p in paths:
            p.unlink(missing_ok=True)
        job.update(status='limit', progress=100, message=str(e))
        write_job(job_id, job)
        return _job_response(job_id)
    except ValueError as e:
        for p in paths:
            p.unlink(missing_ok=True)
        job.update(status='error', progress=100, message=str(e))
        write_job(job_id, job)
        return _job_response(job_id)
    except Exception:
        logger.exception("Could not start job")
        for p in paths:
            p.unlink(missing_ok=True)
        job.update(status='error', progress=100, message="Could not read the uploaded files. Please try again.")
        write_job(job_id, job)
        return _job_response(job_id)

    write_job(job_id, job)
    request.session['active_job'] = job_id
    args = (job_id, mode, [str(p) for p in paths], dict(request.POST.items()), is_pro,
            request.user.pk if request.user.is_authenticated else None)
    if getattr(settings, 'JOBS_RUN_INLINE', False):
        run_job(*args, threaded=False)
    else:
        threading.Thread(target=run_job, args=args, daemon=True, name=f"job-{job_id[:8]}").start()
    return _job_response(job_id)


@require_GET
def job_status(request, job_id):
    job = get_owned_job(request, job_id)
    if not job:
        raise Http404
    return JsonResponse({
        'status': job['status'],
        'progress': job.get('progress', 0),
        'message': job.get('message', ''),
        'result_url': reverse('job_result', args=[job_id]),
    })


@require_GET
def job_result(request, job_id):
    job = get_owned_job(request, job_id)
    if not job:
        return redirect('app')

    context = base_context(request)
    context['active_tab'] = job.get('mode', 'single')
    status = job['status']

    if status in ('queued', 'running'):
        # still working (e.g. the user reloaded the page): the page keeps polling
        context['resume_job'] = {
            'status_url': reverse('job_status', args=[job_id]),
            'result_url': reverse('job_result', args=[job_id]),
        }
    else:
        if request.session.get('active_job') == job_id:
            del request.session['active_job']
        if status == 'done':
            remember_files(request, job.get('files', {}))
            context.update(job.get('context', {}))
        elif status == 'limit':
            context['limit_error'] = job['message']
        else:
            context['error'] = job['message']
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