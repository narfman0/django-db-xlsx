==============
django-db-xlsx
==============

.. image:: https://badge.fury.io/py/django-db-xlsx.png
    :target: https://badge.fury.io/py/django-db-xlsx

.. image:: https://travis-ci.org/narfman0/django-db-xlsx.png?branch=master
    :target: https://travis-ci.org/narfman0/django-db-xlsx

Import/export django models from/to xlsx

Quickstart
----------

Install django-db-xlsx::

    pip install django-db-xlsx

Invoke export::

    from django_db_xlsx import dump_models
    return dump_models()

Invoke import (note: can import exported file)::

    from django_db_xlsx import load_models
    <load workbook from whatever, e.g. FileField in form>
    load_models(wb)

Optional setting for default models saved::

    DEFAULT_DJANGO_DB_XLSX_MODELS = [
        ('myapp', 'mymodel'),
    ]

Features
--------

* Import models via xlsx
* Export models via xlsx

TODO
----

Beef up woefully lacking tests

Running Tests
-------------

Does the code actually work?::

    source <YOURVIRTUALENV>/bin/activate
    (myenv) $ pip install tox
    (myenv) $ tox

License
-------

Copyright Jon Robison 2017, see LICENSE for details
