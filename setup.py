# coding: utf-8
from setuptools import setup
import os


setup(name='django-excel-fixture',
      version='0.1',
      description='Django: import data, export data and test using xlsx files.',
      long_description=open(os.path.join(os.path.dirname(__file__), "README.md")).read(),
      long_description_content_type='text/markdown',
      author="Henrique Portela, SKisContent",
      author_email="henrique.portela@gmail.com",
      license="MIT",
      packages=[
          'django_excel_fixture',
          'django_excel_fixture.serializers',
          ],
      install_requires=[],
      zip_safe=True,
      platforms='any',
      include_package_data=True,
      classifiers=[
          'Development Status :: 5 - Production/Stable',
          'Framework :: Django',
          'Intended Audience :: Developers',
          'License :: OSI Approved :: MIT License',
          'Natural Language :: English',
          'Operating System :: OS Independent',
          'Programming Language :: Python',
          'Programming Language :: Python :: 3',
          'Topic :: Software Development :: Libraries',
      ],
      url='http://github.com/portela/django-excel-fixture',
)