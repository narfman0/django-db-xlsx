#!/usr/bin/env python
# -*- coding: utf-8 -*-

from django.contrib.auth.models import User
from django.http import HttpResponse
from django.test import TestCase

from django_db_xlsx import load_models, dump_models

from openpyxl import Workbook


class TestUtil(TestCase):
    def setUp(self):
        self.user = User.objects.create_user('jon', 'narfl@narfl.rfl', 'jonpass')
        self.target_models = [
            ('auth', 'User'),
        ]

    def tearDown(self):
        self.user.delete()

    def test_dump_load(self):
        wb = Workbook()
        response = dump_models(target_models=self.target_models, wb=wb)
        self.assertTrue(type(response) is HttpResponse)
        # load_models(wb, target_models=self.target_models)
