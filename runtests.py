#!/usr/bin/env python

import os
import sys
import subprocess
from importlib import import_module


if __name__ == '__main__':
    # We need to set the PYTHONPATH environment variable
    # because otherwise subprocesses on Travis CI won't
    # include this directory in the pythonpath
    root_dir = os.path.dirname(os.path.realpath(__file__))
    os.environ['PYTHONPATH'] = root_dir + os.pathsep + os.environ.get('PYTHONPATH', '')

    # Test using django.test.runner.DiscoverRunner
    os.environ['DJANGO_SETTINGS_MODULE'] = 'tests.settings'

    # We need to use subprocess.call instead of django's execute_from_command_line
    # because we can only setup django's settings once, and it's bad
    # practice to change them at runtime
    subprocess.call(['django-admin', 'test', '--nomigrations'])

    # Temp / working / playing around:
    # subprocess.call(['django-admin', 'makemigrations'])
    # subprocess.call(['django-admin', 'migrate'])
    # subprocess.call(['django-admin', 'dumpdata')
    # subprocess.call(['django-admin', 'dumpdata', '--format', 'xlsx', '--output', 'testfile.xlsx'])
    # subprocess.call(['django-admin', 'dumpdata', '--output', 'myfilename.json'])