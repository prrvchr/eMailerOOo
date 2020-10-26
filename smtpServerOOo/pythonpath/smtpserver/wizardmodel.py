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

import validators
import traceback


class WizardModel(unohelper.Base):
    def __init__(self, ctx, email=None):
        self.ctx = ct
        self._User = None
        self._Servers = ()
        self._index = 0
        if email is not None:
            self.Email = email
        self._timeout = self.getTimeout()
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)
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
    def User(self):
        return self._User
    @User.setter
    def User(self, user):
        self._User = user

    @property
    def Servers(self):
        return self._Servers
    @Servers.setter
    def Servers(self, servers):
        self._Servers = servers

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

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

    def isEmailValid(self):
        if validators.email(self.Email):
           return True
        return False

    def getServer(self):
        return self._Servers[self._index].getValue('Server')

    def getServerPage(self):
        count = len(self._Servers)
        if count > 0:
            return self, '%s/%s' % (self._index +1, count)
        else:
            return  None, '1/0'

    def previousServerPage(self):
        self._index -= 1

    def nextServerPage(self):
        self._index += 1

    def isFirst(self):
        return self._index == 0

    def isLast(self):
        return self._index + 1 >= len(self._Servers)

    def getPort(self):
        return self._Servers[self._index].getValue('Port')

    def getConnection(self):
        return self._Servers[self._index].getValue('Connection')

    def getAuthentication(self):
        return self._Servers[self._index].getValue('Authentication')

    def getUserName(self):
        mode = self._Servers[self._index].getValue('LoginMode')
        if mode != -1:
            return self.Email.split('@')[mode]
        return self.Email

    def getSmtpConfig(self, progress, callback):
        self._datasource.getSmtpConfig(self.Email, progress, callback)
