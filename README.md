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

## Plan limits

Defined at the top of `compare_app/views.py`: Free 1,000 rows / 5 MB, Pro 1,048,575 rows / 150 MB.

## Not committed to git

`db.sqlite3`, `processed_files/`, `venv/`, `__pycache__/` and `.env` are in `.gitignore`. Never commit them.
