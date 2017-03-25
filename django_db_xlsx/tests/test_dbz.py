#!/usr/bin/env python
# -*- coding: utf-8 -*-

from django.http import HttpResponse
from django.test import TestCase

from django_db_xlsx import load_models, dump_models


class TestUtil(TestCase):
    def test_dump_load(self):
        response = dump_models(target_models=[])
        self.assertTrue(type(response) is HttpResponse)
        load_models(target_models=[])
