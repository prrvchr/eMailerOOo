#!
# -*- coding: utf_8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020 https://prrvchr.github.io                                     ║
║                                                                                    ║
║   Permission is hereby granted, free of charge, to any person obtaining            ║
║   a copy of this software and associated documentation files (the "Software"),     ║
║   to deal in the Software without restriction, including without limitation        ║
║   the rights to use, copy, modify, merge, publish, distribute, sublicense,         ║
║   and/or sell copies of the Software, and to permit persons to whom the Software   ║
║   is furnished to do so, subject to the following conditions:                      ║
║                                                                                    ║
║   The above copyright notice and this permission notice shall be included in       ║
║   all copies or substantial portions of the Software.                              ║
║                                                                                    ║
║   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,                  ║
║   EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES                  ║
║   OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.        ║
║   IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY             ║
║   CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,             ║
║   TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE       ║
║   OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.                                    ║
║                                                                                    ║
╚════════════════════════════════════════════════════════════════════════════════════╝
"""

import unohelper

from com.sun.star.uno import XCurrentContext
from com.sun.star.mail import XAuthenticator

from smtpserver import KeyMap

from smtpserver import getConfiguration
from smtpserver import getStringResource
from smtpserver import g_identifier
from smtpserver import g_extension

import validators
import json
import traceback


class IspdbModel(unohelper.Base):
    def __init__(self, ctx, datasource, close, email=''):
        self._datasource = datasource
        self._close = close
        self._email = email
        self._user = None
        self._servers = []
        self._metadata = {}
        self._index = 0
        self._count = 0
        self._offline = 0
        self._loaded = False
        self._refresh = False
        self._isnew = False
        self._diposed = False
        self._updated = False
        secure = {0: 3, 1: 4, 2: 4, 3: 5}
        unsecure = {0: 0, 1: 1, 2: 2, 3: 2}
        self._levels = {0: unsecure, 1: secure, 2: secure}
        self._connections = {0: 'Insecure', 1: 'Ssl', 2: 'Tls'}
        self._authentications = {0: 'None', 1: 'Login', 2: 'Login', 3: 'OAuth2'}
        self._configuration = getConfiguration(ctx, g_identifier, True)
        self._timeout = self._configuration.getByName('ConnectTimeout')
        self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._resources = {'Step': 'IspdbPage%s.Step',
                           'Title': 'IspdbPage%s.Title.%s',
                           'Label': 'IspdbPage%s.Label1.Label',
                           'Progress': 'IspdbPage2.Label2.Label.%s',
                           'Security': 'IspdbPage3.Label10.Label.%s',
                           'SendTitle': 'SendDialog.Title',
                           'SendSubject': 'SendDialog.TextField2.Text',
                           'SendMessage': 'SendDialog.TextField3.Text'}

    @property
    def Email(self):
        return self._email
    @Email.setter
    def Email(self, email):
        self._email = email

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

    @property
    def DataSource(self):
        return self._datasource

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

# IspdbModel getter methods
    def saveTimeout(self):
        self._configuration.replaceByName('ConnectTimeout', self._timeout)
        if self._configuration.hasPendingChanges():
            self._configuration.commitChanges()

# IspdbModel getter methods called by IspdbWizard
    def dispose(self):
        self._diposed = True
        if self._close:
            self.DataSource.dispose()

# IspdbModel getter methods called by WizardPages 2 and 4
    def isDisposed(self):
        return self._diposed

# IspdbModel getter methods called by WizardPage1
    def isEmailValid(self, email):
        if validators.email(email):
            return True
        return False

# IspdbModel getter methods called by WizardPage2
    def setUser(self, user):
        self._user = user
        self._metadata['User'] = user.toJson()

    def setServers(self, servers):
        self._isnew = False
        if not len(servers):
            self._isnew = True
            servers = self._getDefaultServers()
        self._metadata['Servers'] = tuple(server.toJson() for server in servers)
        self._servers = servers
        self._index = self._getServerIndex(0)

    def getSmtpConfig(self, *args):
        self._refresh = True
        self._loaded = False
        url = self._configuration.getByName('IspDBUrl')
        self._datasource.getSmtpConfig(self.Email, url, *args)

# IspdbModel getter methods called by WizardPage3
    def isRefreshed(self):
        if self._refresh:
            self._refresh = False
            self._loaded = True
            return True
        return False

    def getConfig(self):
        config = KeyMap()
        config.setValue('First', self._isFirst())
        config.setValue('Last', self._isLast())
        config.setValue('Page', self._getServerPage())
        config.setValue('Default', self._isDefaultPage())
        config.update(self._getServer())
        config.setValue('Login', self._getLoginName())
        config.setValue('Password', self._getPassword())
        return config

    def getAuthentication(self):
        return self._getServer().getValue('Authentication')

    def isConnectionValid(self, host, port):
        return self._isHostValid(host) and self._isPortValid(port)

    def isStringValid(self, value):
        if validators.length(value, min=1):
            return True
        return False

    def previousServerPage(self, server):
        self._getServer().update(server)
        self._index -= 1

    def nextServerPage(self, server):
        self._getServer().update(server)
        self._index += 1

    def updateConfiguration(self, user, server):
        print("PageModel.updateConfiguration() %s - %s" % (user, server))
        self._getServer().update(server)
        print("PageModel.updateConfiguration() server:\n%s\n%s" % (server.toJson(), self._getServerMetaData()))
        self._user.update(user)
        print("PageModel.updateConfiguration() user:\n%s\n%s" % (user.toJson(), self._getUserMetaData()))

    def saveConfiguration(self):
        server = self._getServer()
        if self._isnew or server.toJson() != self._getServerMetaData():
            provider = self._getDomain()
            host, port = self._getServerKeys()
            self._datasource.saveServer(self._isnew, provider, host, port, server)
            print("PageModel.saveConfiguration() server:\n%s\n%s" % (server.toJson(), self._metadata['Servers'][self._index]))
        if self._user.toJson() != self._getUserMetaData():
            print("PageModel.saveConfiguration() user:\n%s\n%s" % (self._user.toJson(), self._getUserMetaData()))
            self._datasource.saveUser(self.Email, self._user)

    def getSecurity(self, i, j):
        level = self._levels.get(i).get(j)
        message = self.getSecurityMessage(level)
        return message, level

# IspdbModel getter methods called by WizardPage4
    def getHost(self):
        return self._getServer().getValue('Server')

    def getPort(self):
        return self._getServer().getValue('Port')

    def smtpConnect(self, *args):
        context = self._getConnectionContext()
        authenticator = self._getAuthenticator()
        self._datasource.smtpConnect(context, authenticator, *args)

    def smtpSend(self, *args):
        context = self._getConnectionContext()
        authenticator = self._getAuthenticator()
        self._datasource.smtpSend(context, authenticator, self.Email, *args)

# IspdbModel getter methods called by SendDialog
    def validSend(self, recipient, subject, message):
        enabled = all((self.isEmailValid(recipient),
                       self.isStringValid(subject),
                       self.isStringValid(message)))
        return enabled

# IspdbModel private getter methods
    def _getServer(self):
        return self._servers[self._index]

    def _getDomain(self):
        return self._user.getValue('Domain')

# IspdbModel private getter methods called by WizardPage2
    def _getDefaultServers(self):
        server = KeyMap()
        server.setValue('Server', 'smtp.%s' % self._getDomain())
        server.setValue('Port', 25)
        server.setValue('Connection', 0)
        server.setValue('Authentication', 0)
        server.setValue('LoginMode', 1)
        return [server, ]

    def _getServerIndex(self, default=-1):
        port = self._user.getValue('Port')
        server = self._user.getValue('Server')
        if port != 0:
            for s in self._servers:
                if s.getValue('Server') == server and s.getValue('Port') == port:
                    default = self._servers.index(s)
                    break;
        return default

# IspdbModel private getter methods called by WizardPage3
    def _isFirst(self):
        return self._index == 0

    def _isLast(self):
        count = len(self._servers)
        return self._index +1 >= count

    def _getServerPage(self):
        if self._isnew:
            page = '1/0'
        else:
            count = len(self._servers)
            page = '%s/%s' % (self._index +1, count)
        return page

    def _isDefaultPage(self):
        return self._index == self._getServerIndex()

    def _getLoginName(self):
        login = self._user.getValue('LoginName')
        return login if login != '' else self._getLoginFromEmail()

    def _getLoginFromEmail(self):
        mode = self._getLoginMode()
        return self.Email.partition('@')[mode] if mode != 1 else self.Email

    def _getLoginMode(self):
        return self._getServer().getValue('LoginMode')

    def _getPassword(self):
        return self._user.getValue('Password')

    def _isHostValid(self, host):
        if validators.domain(host):
            return True
        return False

    def _isPortValid(self, port):
        if validators.between(port, min=1, max=1023):
            return True
        return False

    def _getUserMetaData(self):
        return self._metadata['User']

    def _getServerMetaData(self):
        return self._metadata['Servers'][self._index]

    def _getServerKeys(self):
        server = json.loads(self._getServerMetaData())
        return server['Server'], server['Port']

# IspdbModel private getter methods called by WizardPage4
    def _getConnectionContext(self):
        server = self._getServer().getValue('Server')
        port = self._getServer().getValue('Port')
        connection = self._getConnectionMap()
        authentication = self._getAuthenticationMap()
        data = {'ServerName': server,
                'Port': port,
                'ConnectionType': connection,
                'AuthenticationType': authentication,
                'Timeout': self.Timeout}
        return CurrentContext(data)

    def _getAuthenticator(self):
        user = self._getLoginName()
        password = self._getPassword()
        return Authenticator(user, password)

    def _getConnectionMap(self):
        index = self._getServer().getValue('Connection')
        return self._connections.get(index)

    def _getAuthenticationMap(self):
        index = self._getServer().getValue('Authentication')
        return self._authentications.get(index)

# IspdbModel StringResource methods
    def getPageStep(self, pageid):
        resource = self._resources.get('Step') % pageid
        return self._resolver.resolveString(resource)

    def getPageTitle(self, pageid):
        resource = self._resources.get('Title') % (pageid, self._offline)
        return self._resolver.resolveString(resource)

    def getPageLabel(self, pageid):
        resource = self._resources.get('Label') % pageid
        return self._resolver.resolveString(resource)

    def getProgressMessage(self, value):
        resource = self._resources.get('Progress') % value
        return self._resolver.resolveString(resource)

    def getSecurityMessage(self, level):
        resource = self._resources.get('Security') % level
        return self._resolver.resolveString(resource)

    def getSendTitle(self):
        resource = self._resources.get('SendTitle')
        return self._resolver.resolveString(resource)

    def getSendSubject(self):
        resource = self._resources.get('SendSubject')
        return self._resolver.resolveString(resource)

    def getSendMessage(self):
        resource = self._resources.get('SendMessage')
        return self._resolver.resolveString(resource)


class CurrentContext(unohelper.Base,
                     XCurrentContext):
    def __init__(self, context):
        self._context = context

    def getValueByName(self, name):
        return self._context[name]


class Authenticator(unohelper.Base,
                    XAuthenticator):
    def __init__(self, user, password):
        self._user = user
        self._password = password

    def getUserName(self):
        return self._user

    def getPassword(self):
        return self._password
