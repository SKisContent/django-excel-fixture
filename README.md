# Django Excel Fixture

Django Fixtures using excel files. Easily importing, exporting and test data.

[![Build Status](https://travis-ci.org/portela/django-excel-fixture.svg?branch=master)](https://travis-ci.org/portela/django-excel-fixture)


## Installation
-
Soon, to be pip-ified project. For now, install by hand (see master branch).


## Usage
-
If you already have some data in the database, to save it use:

```./manage.py dumpdata myapp1.MyModelName --format xlsx --output myfilename.xlsx```
 
 If you leave out the **output** option, the table is dumped to stdout in CSV format.

To manually load the data from the Excel file, use:

```./manage.py loaddata myfilename.xlsx```
 
To use the fixture in a test, just add

```    fixtures = ['myfilename.xlsx']```

to the test class.

If you need to brush up on fixtures, look at https://django-testing-docs.readthedocs.io/en/latest/fixtures.html.


## Contribution

```console
git clone git@github.com:SKisContent/django-excel-fixture.git
cd django-excel-fixture
python -m venv .venv
source .venv/scripts/activate
pip install -r requirements.txt
python runtests.py
```

 **_NOTE:_**  On Windows, change line 4 to


```console
    cd ./.venv/scripts/
    activate
```


Authors
------

* [Henrique Portela](https://github.com/portela)
* [SKisContent](https://github.com/SKisContent)


License
=======

The MIT License.