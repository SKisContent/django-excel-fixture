# coding: utf-8

from django.conf import settings
from django.core import serializers

from django.test import TestCase
from openpyxl import Workbook
from unittest import skip

from .models import Person, Recipe, Ingredient, Association


class XlsxSerializerTest(TestCase):
    fixtures = ['fixture.json']

    @classmethod
    def setUpTestData(cls):
        pass

    def setUp(self):
        pass

    def test_serializer_has_been_registered(self):
        public_formats = serializers.get_public_serializer_formats()
        self.assertIn('xlsx', public_formats)

    def _get_field_values(self, stream, field_name):
        pass

    def _get_pk_value(self, stream):
        pass

    def _validate_output(self, stream):
        pass


class XlsxSerializerUnitTest(TestCase):
    fixtures = ['fixture.json']

    def setUp(self):
        XLSXSerializer = serializers.get_serializer("xlsx")
        self.xlsx_serializer = XLSXSerializer()

    def test_start_serialization(self):
        self.assertIsNone(self.xlsx_serializer.workbook)
        self.xlsx_serializer.start_serialization()
        self.assertIsInstance(self.xlsx_serializer.workbook, Workbook)

    def test_after_start_number_of_sheet(self):
        self.xlsx_serializer.start_serialization()
        self.assertEquals(0, len(self.xlsx_serializer.workbook.worksheets))

    def _start_object(self):
        self.obj = Person(name='Henrique Portela', age=41)
        self.xlsx_serializer.start_serialization()
        self.xlsx_serializer.start_object(self.obj)

    def test_start_object_create_sheet(self):
        """ If the obj sheet doesn't exist, it must be created. """
        self._start_object()
        self.assertEquals(1, len(self.xlsx_serializer.workbook.worksheets))

    def test_start_object_sheet_name(self):
        """ Name of the sheet, should be obj class. """
        self._start_object()
        self.assertIn('myapp.Person', self.xlsx_serializer.workbook.sheetnames)

    def test_start_object_sheet_header(self):
        """ Newly created sheet, must have header. """
        self._start_object()
        self.assertEquals('id' , self.xlsx_serializer.workbook['myapp.Person']['A1'].value)
        self.assertEquals('name' , self.xlsx_serializer.workbook['myapp.Person']['B1'].value)
        self.assertEquals('age' , self.xlsx_serializer.workbook['myapp.Person']['C1'].value)

    def test_sheet_header_format(self):
        """ The header should be formated. """
        self._start_object()
        self.assertEquals('Calibri', self.xlsx_serializer.workbook['myapp.Person']['A1'].font.name)
        self.assertEquals(11, self.xlsx_serializer.workbook['myapp.Person']['A1'].font.size)
        self.assertEquals(True, self.xlsx_serializer.workbook['myapp.Person']['A1'].font.bold)
        self.assertEquals('solid', self.xlsx_serializer.workbook['myapp.Person']['A1'].fill.patternType)
        self.assertEquals('009BBB59', self.xlsx_serializer.workbook['myapp.Person']['A1'].fill.fgColor.rgb)
        self.assertEquals('center', self.xlsx_serializer.workbook['myapp.Person']['A1'].alignment.horizontal)
        self.assertEquals('center', self.xlsx_serializer.workbook['myapp.Person']['A1'].alignment.vertical)

    def test_start_object_current_row(self):
        """
        After creating a new sheet, current row should be pointing to line 2.
        """
        self._start_object()
        self.assertEquals(2, self.xlsx_serializer.current_row)

    def test_start_object_curret_row_update(self):
        """ After start object, current_row should be incremented. """
        self._start_object()
        for i in range(3,10):
            self.obj = Person(name='Henrique Portela', age=41)
            self.xlsx_serializer.start_object(self.obj)
            self.assertEquals(i, self.xlsx_serializer.current_row)

    def test_start_object_active_sheet(self):
        """ When second obj is started, active sheet should change"""
        self._start_object()
        self.obj = Association.objects.first()
        self.xlsx_serializer.start_object(self.obj)
        last_sheet = self.xlsx_serializer.workbook['myapp.Association']
        self.assertEquals(last_sheet, self.xlsx_serializer.workbook.active)

    def test_handle_field(self):
        self._start_object()

        self.xlsx_serializer.handle_field(self.obj, self.obj._meta.fields[1])
        self.assertEquals('Henrique Portela', self.xlsx_serializer.workbook['myapp.Person']['B2'].value)
        self.xlsx_serializer.handle_field(self.obj, self.obj._meta.fields[2])
        self.assertEquals(41, self.xlsx_serializer.workbook['myapp.Person']['C2'].value)

        self.obj = Person(name='John Smith', age=51)
        self.xlsx_serializer.start_object(self.obj)

        self.xlsx_serializer.handle_field(self.obj, self.obj._meta.fields[1])
        self.assertEquals('John Smith', self.xlsx_serializer.workbook['myapp.Person']['B3'].value)
        self.xlsx_serializer.handle_field(self.obj, self.obj._meta.fields[2])
        self.assertEquals(51, self.xlsx_serializer.workbook['myapp.Person']['C3'].value)

    @skip('DateTimeField without value, should now crash')
    def test_handle_field_empty(self):
        pass

    def test_handle_fk_field(self):
        self.obj = Association.objects.first()
        self.xlsx_serializer.start_serialization()
        self.xlsx_serializer.start_object(self.obj)

        self.xlsx_serializer.handle_fk_field(self.obj, self.obj._meta.fields[1])
        self.assertEquals(1, self.xlsx_serializer.workbook['myapp.Association']['B2'].value)
        self.xlsx_serializer.handle_fk_field(self.obj, self.obj._meta.fields[2])
        self.assertEquals(1, self.xlsx_serializer.workbook['myapp.Association']['C2'].value)