#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import KeyMap
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


class PageModel(unohelper.Base):
    def __init__(self, ctx, email=None):
        self.ctx = ctx
        self._User = None
        self._Servers = ()
        self._index = self._default = self._count = 0
        self._isnew = False
        if email is not None:
            self.Email = email
        self._timeout = self.getTimeout()
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)
        try:
            msg = "PageModel.__init__()"
            print(msg)
            self._datasource = DataSource(self.ctx)
        except Exception as e:
            msg = "PageModel.__init__(): Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    _Email = ''

    @property
    def Email(self):
        return PageModel._Email
    @Email.setter
    def Email(self, email):
        PageModel._Email = email

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
        self._Servers = self._getServers(servers)

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

    def isServerValid(self, server):
        if validators.domain(server):
           return True
        return False

    def isPortValid(self, port):
        if validators.between(port, min=1, max=1023):
           return True
        return False

    def isStringValid(self, value):
        if validators.length(value, min=1):
           return True
        return False

    def getServer(self):
        return self._Servers[self._index].getValue('Server')

    def getPort(self):
        return self._Servers[self._index].getValue('Port')

    def getConnection(self):
        return self._Servers[self._index].getValue('Connection')

    def getAuthentication(self):
        return self._Servers[self._index].getValue('Authentication')

    def getLoginName(self):
        login = self._User.getValue('LoginName')
        if login != '':
            return login
        mode = self._Servers[self._index].getValue('LoginMode')
        if mode != -1:
            return self.Email.split('@')[mode]
        return self.Email

    def getPassword(self):
        return self._User.getValue('Password')

    def getServerPage(self):
        if self._isnew:
            return '1/0'
        return '%s/%s' % (self._index + 1, self._count)

    def previousServerPage(self):
        self._index -= 1

    def nextServerPage(self):
        self._index += 1

    def isFirst(self):
        return self._index == 0

    def isLast(self):
        return self._index + 1 >= self._count

    def getSmtpConfig(self, progress, callback):
        self._datasource.getSmtpConfig(self.Email, progress, callback)

    def _getServers(self, servers):
        self._isnew = False
        self._count = len(servers)
        self._index = self._default = 0
        if self._count == 0:
            self._count = 1
            self._isnew = True
            servers = self._getDefaultServers()
        elif self._count > 1:
            self._index = self._default = self._getServerIndex(servers)
        return servers

    def _getDefaultServers(self):
        server = KeyMap()
        server.setValue('Server', '')
        server.setValue('Port', 25)
        server.setValue('Connection', 0)
        server.setValue('Authentication', 0)
        server.setValue('LoginMode', -1)
        return (server, )

    def _getServerIndex(self, servers):
        index = 0
        port = self.User.getValue('Port')
        if port != 0:
            for server in servers:
                if server.getValue('Port') == port:
                    index = servers.index(server)
                    break;
        return index
