#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService
from unolib import getConfiguration
from unolib import getStringResource

from .replicator import Replicator

from .configuration import g_identifier
from .configuration import g_extension

from .logger import logMessage
from .logger import getMessage

import traceback


class WizardModel(unohelper.Base):
    def __init__(self, ctx, email=''):
        self.ctx = ctx
        self._email = email
        self._timeout = self.getTimeout()
        try:
            msg = "WizardModel.__init__()"
            print(msg)
            #self._replicator = Replicator(self.ctx)
        except Exception as e:
            msg = "WizardModel.__init__(): Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    @property
    def Email(self):
        return self._email
    @Email.setter
    def Email(self, email):
        self._email = email

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
