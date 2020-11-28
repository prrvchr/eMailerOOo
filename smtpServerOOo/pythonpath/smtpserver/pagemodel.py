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
import json
import traceback


class PageModel(unohelper.Base):
    def __init__(self, ctx, email=None):
        self.ctx = ctx
        self._User = None
        self._Servers = ()
        self._metadata = ()
        self._index = self._default = self._count = self._offline = 0
        self._isnew = False
        if email is not None:
            self.Email = email
        self._timeout = self.getTimeout()
        self._connections = {0: 'Insecure', 1: 'Ssl', 2: 'Tls'}
        self._authentications = {0: 'None', 1: 'Login', 2: 'Login', 3: 'OAuth2'}
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
        self._count = len(servers)
        if self._count == 0:
            self._count = 1
            self._isnew = True
            self._index = 0
            self._default = -1
            servers = self._getDefaultServers()
        else:
            self._isnew = False
            self._index, self._default = self._getServerIndex(servers)
        self._metadata = tuple(s.toJson() for s in servers)
        self._Servers = servers

    @property
    def Offline(self):
        return self._offline
    @Offline.setter
    def Offline(self, offline):
        self._offline = offline

    @property
    def Timeout(self):
        return self._timeout
    @Timeout.setter
    def Timeout(self, timeout):
        self._timeout = timeout

    def getConnectionMap(self, index):
        return self._connections.get(index)

    def getAuthenticationMap(self, index):
        return self._authentications.get(index)

    def getTimeout(self):
        return getConfiguration(self.ctx, g_identifier, False).getByName('ConnectTimeout')
    def saveTimeout(self):
        configuration = getConfiguration(self.ctx, g_identifier, True)
        configuration.replaceByName('ConnectTimeout', self._timeout)
        if configuration.hasPendingChanges():
            configuration.commitChanges()

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

    def isEmailValid(self, email):
        if validators.email(email):
            return True
        return False

    def isHostValid(self, host):
        if validators.domain(host):
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

    def getHost(self):
        return self._getServer().getValue('Server')

    def getPort(self):
        return self._getServer().getValue('Port')

    def getConnection(self):
        return self._getServer().getValue('Connection')

    def getAuthentication(self):
        return self._getServer().getValue('Authentication')

    def getLoginName(self):
        login = self._User.getValue('LoginName')
        return login if login != '' else self._getLoginFromEmail()

    def getLoginMode(self):
        return self._getServer().getValue('LoginMode')

    def getPassword(self):
        return self._User.getValue('Password')

    def getServerPage(self):
        if self._isnew:
            return '1/0'
        return '%s/%s' % (self._index + 1, self._count)

    def previousServerPage(self, server):
        self._getServer().update(server)
        self._index -= 1

    def nextServerPage(self, server):
        self._getServer().update(server)
        self._index += 1

    def isFirst(self):
        return self._index == 0

    def isLast(self):
        return self._index + 1 >= self._count

    def getSmtpConfig(self, *args):
        self._datasource.getSmtpConfig(self.Email, *args)

    def smtpConnect(self, *args):
        self._datasource.smtpConnect(*args)

    def smtpSend(self, *args):
        self._datasource.smtpSend(*args)

    def saveConfiguration(self, user, server):
        if self._isnew or server.toJson() != self._metadata[self._index]:
            provider = user.getValue('Domain')
            host, port = self._getServerKeys()
            self._datasource.saveServer(self._isnew, provider, host, port, server)
            print("PageModel.saveConfiguration() server:\n%s\n%s" % (server.toJson(), self._metadata[self._index]))
        if user.toJson() != self.User.toJson():
            print("PageModel.saveConfiguration() user:\n%s\n%s" % (user.toJson(), self.User.toJson()))
            self._datasource.saveUser(self.Email, user)

    def _getLoginFromEmail(self):
        mode = self.getLoginMode()
        return self.Email.partition('@')[mode] if mode != 1 else self.Email

    def _getServerKeys(self):
        server = json.loads(self._metadata[self._index])
        return server['Server'], server['Port']

    def _getServer(self):
        return self._Servers[self._index]

    def _getDefaultServers(self):
        server = KeyMap()
        server.setValue('Server', '')
        server.setValue('Port', 25)
        server.setValue('Connection', 0)
        server.setValue('Authentication', 0)
        server.setValue('LoginMode', 1)
        return (server, )

    def _getServerIndex(self, servers):
        index = 0
        default = -1
        port = self.User.getValue('Port')
        if port != 0:
            for server in servers:
                if server.getValue('Port') == port:
                    index = default = servers.index(server)
                    break;
        return index, default
