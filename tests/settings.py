# coding: utf-8


ALLOWED_HOSTS = ['*']



# Temporary adding sqllite to work / test / play
# with serialize
import os
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        # Just for now:
        'NAME': os.path.join(BASE_DIR, 'db.sqlite3'),
        #'NAME': ':memory:',
    }
}

INSTALLED_APPS = (
    'tests.myapp',
    'test_without_migrations',
    'django_excel_fixture',
    'django_extensions',
)

SITE_ID=1,

SECRET_KEY='secret'

MIDDLEWARE_CLASSES = (
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
)