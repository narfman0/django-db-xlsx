#!/usr/bin/env python
# -*- coding: utf-8 -*-
try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

version = '1.0.1'
readme = open('README.rst').read()

setup(
    name='django-db-xlsx',
    version=version,
    description="Import/export django models via xlsx",
    long_description=readme,
    author='Jon Robison',
    author_email='narfman0@gmail.com',
    url='https://github.com/narfman0/django-db-xlsl',
    packages=[
        'django_db_xlsx',
    ],
    include_package_data=True,
    install_requires=[
        'django',
        'openpyxl',
    ],
    license="MIT",
    zip_safe=False,
    keywords=['django', 'database', 'xlsx'],
    classifiers=[
        'Development Status :: 4 - Beta',
        'Framework :: Django',
        'Framework :: Django :: 1.8',
        'Framework :: Django :: 1.9',
        'Framework :: Django :: 1.10',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Natural Language :: English',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.5',
    ],
)
