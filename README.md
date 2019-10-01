# Django Excel Fixture

[![Build Status](https://travis-ci.org/portela/django-excel-fixture.svg?branch=master)](https://travis-ci.org/portela/django-excel-fixture)

Django-excel-fixture is an easy way to import, export and test using .xlsx file.

## Motivation

Excel is a must in the corporate world. It is an easy to use and highly powerful application, currently being used by millions of users (maybe more than a billion user). Django-excel-fixture provides a way to import, export data using excel file (.xlsx); and also use it as test fixture.


## Installation

To install it, simply:

```console
$ pip install django-excel-fixture
```

Then add it to your INSTALLED_APPS on your settings.py:

```python
INSTALLED_APPS = (
    #...
    'django_excel_fixture',
)
```

## Usage


### Exporting data from database:

If you already have some data in the database, to export it to a .xlsx file:

```console
$ python manage.py dumpdata --format xlsx --output myfilename.xlsx
```

It is also possible to select a specific model:

```console
$ python manage.py dumpdata myapp1.MyModelName --format xlsx --output myfilename.xlsx
```

* If you leave out the **output** option, the table is dumped to stdout in CSV format.
* There will be sheets only for table with data

### Importing data into the database:

To manually load the data from the Excel file, use:

```console
$ python manage.py loaddata myfilename.xlsx
```

### Using as test fixture:

To use the fixture in a test, just add to the test class:

```python
class MyTestCase(TestCase):
    fixtures = ['myfilename.xlsx']

    def ...
```

* No performance optimization has been made. Caution is advised.


## Excel file format

* Currently, it is only compatible with .xlxs format.
* Every sheet is a model. The sheet title represents the model's identification.
* Every column represents a field.
* Every row represents a object (database entry).


## Contribution

```console
git clone git@github.com:SKisContent/django-excel-fixture.git
cd django-excel-fixture
python -m venv .venv
source .venv/scripts/activate
pip install -r requirements.txt
python runtests.py
```

Authors
------

* [Henrique Portela](https://github.com/portela)
* [SKisContent](https://github.com/SKisContent)

License
=======

The MIT License.