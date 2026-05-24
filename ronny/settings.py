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
]

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
