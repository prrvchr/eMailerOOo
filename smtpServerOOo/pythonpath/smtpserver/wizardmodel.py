#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService
from unolib import getConfiguration
from unolib import getStringResource

from .datasource import DataSource

from .configuration import g_identifier
from .configuration import g_extension

from .logger import logMessage
from .logger import getMessage

import traceback


class WizardModel(unohelper.Base):
    def __init__(self, ctx, email=None):
        self.ctx = ctx
        if email is not None:
            self.Email = email
        self._timeout = self.getTimeout()
        try:
            msg = "WizardModel.__init__()"
            print(msg)
            self._datasource = DataSource(self.ctx)
        except Exception as e:
            msg = "WizardModel.__init__(): Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    _Email = ''

    @property
    def Email(self):
        return WizardModel._Email
    @Email.setter
    def Email(self, email):
        WizardModel._Email = email

    @property
    def Timeout(self):
        return self._timeout
    @Timeout.setter
    def Timeout(self, timeout):
        self._timeout = timeout

    def getTimeout(self):
        return getConfiguration(self.ctx, g_identifier, False).getByName('ConnectTimeout')
    def saveTimeout(self):
        configuration = getConfiguration(self.ctx, g_identifier, True)
        configuration.replaceByName('ConnectTimeout', self._timeout)
        if configuration.hasPendingChanges():
            configuration.commitChanges()
