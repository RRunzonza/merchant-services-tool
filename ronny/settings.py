import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

# In production set this environment variable to a long random string.
# Locally it falls back to the dev key so nothing breaks during development.
SECRET_KEY = os.environ.get(
    'SECRET_KEY',
    'django-insecure-change-me-in-production-merchant-services-2024',
)

DEBUG = os.environ.get('DEBUG', 'False') != 'False'

ALLOWED_HOSTS_ENV = os.environ.get('ALLOWED_HOSTS', '*')
ALLOWED_HOSTS = [h.strip() for h in ALLOWED_HOSTS_ENV.split(',') if h.strip()]

INSTALLED_APPS = [
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.staticfiles',
    'merchant',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
    # Automatically removes temp files for expired/orphaned sessions (runs ~hourly)
    'merchant.middleware.TempCleanupMiddleware',
]

# How often (seconds) TempCleanupMiddleware sweeps for expired session temp data.
# Default: 3600 (1 hour).  Override via environment variable if needed.
TEMP_CLEANUP_INTERVAL = int(os.environ.get('TEMP_CLEANUP_INTERVAL', 3600))

ROOT_URLCONF = 'ronny.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [BASE_DIR / 'ronny' / 'templates'],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
            ],
        },
    },
]

WSGI_APPLICATION = 'ronny.wsgi.application'

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}

SESSION_ENGINE = 'django.contrib.sessions.backends.file'
SESSION_COOKIE_AGE = 86400  # 24 hours

# ---------------------------------------------------------------------------
# In-memory cache — DataFrames and generated Excel reports are stored here
# instead of on disk.  This keeps the PythonAnywhere 512 MB quota free for
# the virtualenv and source code.
#
# LocMemCache is per-process (no shared state between gunicorn workers on
# paid plans, but fine for the free single-worker tier).  Objects expire
# automatically after SESSION_COOKIE_AGE seconds (24 h).
# ---------------------------------------------------------------------------
CACHES = {
    'default': {
        'BACKEND': 'django.core.cache.backends.locmem.LocMemCache',
        'LOCATION': 'merchant-services-cache',
    }
}

STATIC_URL = '/static/'

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

TEMP_DATA_DIR = Path(os.environ.get('TEMP_DATA_DIR', str(BASE_DIR / 'temp_data')))

# Create required directories on startup so the app doesn't crash on a fresh server
for _d in [
    TEMP_DATA_DIR / 'sessions',
    TEMP_DATA_DIR / 'uploads',
    TEMP_DATA_DIR / 'pickles',
]:
    _d.mkdir(parents=True, exist_ok=True)

SESSION_FILE_PATH = TEMP_DATA_DIR / 'sessions'

# ---------------------------------------------------------------------------
# Startup cleanup: remove upload/pickle dirs for expired or missing sessions.
# Runs once each time the Django process starts (e.g. after a PythonAnywhere
# reload), so stale files are never left to consume disk quota indefinitely.
# ---------------------------------------------------------------------------
def _startup_cleanup():
    import shutil as _shutil
    import time as _time

    _max_age = SESSION_COOKIE_AGE
    _now = _time.time()
    _sessions_dir = SESSION_FILE_PATH

    for _sub in ('uploads', 'pickles'):
        _base = TEMP_DATA_DIR / _sub
        if not _base.exists():
            continue
        for _sdir in _base.iterdir():
            if not _sdir.is_dir():
                continue
            _sf = _sessions_dir / _sdir.name
            try:
                if not _sf.exists():
                    _shutil.rmtree(_sdir, ignore_errors=True)
                elif _now - _sf.stat().st_mtime > _max_age:
                    _shutil.rmtree(_sdir, ignore_errors=True)
                    _sf.unlink(missing_ok=True)
            except OSError:
                pass

try:
    _startup_cleanup()
except Exception:
    pass  # Never let cleanup crash the server startup
