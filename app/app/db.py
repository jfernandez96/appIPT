from pathlib import Path

# Build paths inside the project like this: BASE_DIR / 'subdir'.

POSGRESQL = {
    'default': {
        'ENGINE': 'django.db.backends.postgresql_psycopg2',
        'NAME': 'MultiVendor_db',
        'USER': 'postgres',
        'PASSWORD': 'adm',
        'HOST': 'localhost',
        'PORT': '5432'
    }
}