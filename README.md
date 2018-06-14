# Django Excel Fixture
I was getting ready to write some tests and needed test data, and it seemed easiest to create it using Excel. However, I couldn't find existing code that would quickly load the data into the db, so I wrote a fixture for loading and saving data to Excel files.

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
This uses the openpyxl package, so you will need to install that.

```pip install openpyxl```


Usage
-
If you already have some data in the database, to save it use:

```./manage.py dumpdata myapp1.MyModelName --format xlsx --output myfilename.xlsx```
 
 If you leave out the **output** option, the table is dumped to <stdout> in CSV format.

To manually load the data from the Excel file, use:

```./manage.py loaddata myfilename.xlsx```
 
To use the fixture in a test, just add

```    fixtures = ['myfilename.xlsx']```

to the test class.

Contribute
-
This gets the job done but there could be more work done to make it more resilient and handle more data types. It also doesn't do anything with foreign keys or many-to-many relationships. If you feel like contributing, please do! 