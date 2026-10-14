# DataFlow Toolkit (Excel-Extractor)

Clean, deduplicate and compare Excel / CSV data. Django app with free and Pro plans.

## Run locally (Windows)

```powershell
git clone https://github.com/souravdas007123/Excel-Extractor.git
cd Excel-Extractor
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python manage.py migrate
python manage.py createsuperuser
python manage.py runserver
```

Open http://127.0.0.1:8000/ (admin: http://127.0.0.1:8000/admin/).

To make a user Pro: admin -> User Profiles -> tick `is_pro` (leave `pro_expires_on` empty for no expiry).

## Tests

```powershell
python manage.py test
```

`compare_app/test_cleaning.py` tests the cleaning logic (pandas only); `compare_app/tests.py` tests the views.

## Settings (environment variables)

| Variable | Default | Notes |
|---|---|---|
| `DJANGO_DEBUG` | `1` | Set `0` in production |
| `DJANGO_SECRET_KEY` | dev key | **Required** when `DJANGO_DEBUG=0` |
| `DJANGO_ALLOWED_HOSTS` | `localhost,127.0.0.1` | Comma separated |
| `SUPPORT_EMAIL` | `support@socialflylive.com` | Shown on landing / upgrade pages |
| `WHATSAPP_NUMBER` | empty | e.g. `919876543210`; empty hides the WhatsApp button |
| `MAX_CONCURRENT_JOBS` | `2` | Files processed at the same time per server process (the rest wait in a queue) |

## Plan limits

Defined at the top of `compare_app/views.py`: Free 1,000 rows / 5 MB, Pro 1,048,575 rows / 150 MB.

## How processing works

Uploads are sent in the background with a real upload progress bar. The server answers immediately with a job id,
processes the file in a background thread and the page shows real progress, then opens the report. Status is kept
in small files under `processed_files/jobs/`, so no Redis or Celery is needed (it also works on Windows). If the page
is reloaded while a file is processing, the progress continues. Without JavaScript the old normal form post still works.

## Production notes

```bash
# Several workers are fine (job status is stored on disk). Do not use --max-requests: it can
# restart a worker while a background job is running.
gunicorn excel_project.wsgi:application --workers 2 --timeout 900
```

Nginx: raise the upload size, and compress text responses (not needed for .xlsx downloads):

```nginx
client_max_body_size 160M;
proxy_read_timeout 900;
gzip on;
gzip_types text/css application/javascript application/json text/plain;
```

Run `python manage.py collectstatic` / serve `/static/` from Nginx if you add static files later.

## Not committed to git

`db.sqlite3`, `processed_files/`, `venv/`, `__pycache__/` and `.env` are in `.gitignore`. Never commit them.
