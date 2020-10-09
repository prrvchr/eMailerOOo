#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService
from unolib import getConfiguration
from unolib import getStringResource

from .configuration import g_identifier
from .configuration import g_extension

from .logger import logMessage
from .logger import getMessage

import traceback


class WizardModel(unohelper.Base):
    def __init__(self, ctx, email=''):
        self.ctx = ctx
        self._email = email

    @property
    def Email(self):
        return self._email
    @Email.setter
    def Email(self, email):
        self._email = email
