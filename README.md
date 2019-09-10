# Django Excel Fixture

Using excel file format for fixture in Django. Easely importanting, exporting and test data.

[![Build Status](https://travis-ci.org/portela/django-excel-fixture.svg?branch=master)](https://travis-ci.org/portela/django-excel-fixture)

Installation
-
At the moment this is not a pip-ified project, so the installation is going to be manual. Copy the contents of the _src_ directory into your django application under the root source directory or one of the apps, although I prefer to create a _core_ package.

*Example*:

If your Django tree looks something like:
```
MyProject
|-myproject
| |-__init__.py
| |-myapp1
| | |-__init__.py
| | |-apps.py
| | |-models.py
| | |-views.py
| |-myapp2
|-templates
```

then post installation it would look like:
```
MyProject
|-myproject
| |-__init__.py
| |-core                  <--
| | |-__init__.py         <--
| | |-serializers         <--
| | | |-__init__.py       <--
| | | |-excel.py          <--
| |-myapp1
| | |-__init__.py
| | |-apps.py
| | |-models.py
| | |-views.py
| |-myapp2
|-templates
```

Dependency
-
- Only tested on Python 3.6

- This uses the openpyxl package, so you will need to install that.

```pip install openpyxl```


Usage
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

## Contribute

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



> **_NOTE:_** This gets the job done but there could be more work done to make it more resilient and handle more data types. It also doesn't do anything with foreign keys or many-to-many relationships. If you feel like contributing, please do!

Author
------

* [Henrique Portela] (https://github.com/portela>)
* [SKisContent](https://github.com/SKisContent>)


License
=======

The MIT License.