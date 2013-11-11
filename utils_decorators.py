# -*- coding: utf-8 -*-
from django.conf import settings


def ignore_this_test(func):
    def wrapper(func):
        func.__test__ = False
        return func
    return wrapper
	