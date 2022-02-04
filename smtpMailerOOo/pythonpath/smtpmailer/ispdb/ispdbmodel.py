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

from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.mail.MailServiceType import IMAP

from smtpmailer import KeyMap
from smtpmailer import DataParser

from smtpmailer import CurrentContext
from smtpmailer import Authenticator

from smtpmailer import getConfiguration
from smtpmailer import getStringResource
from smtpmailer import g_identifier
from smtpmailer import g_extension

from .pages import IspdbServer
from .pages import IspdbUser

import validators
import json
import traceback


class IspdbModel(unohelper.Base):
    def __init__(self, ctx, datasource, close, email=''):
        self._datasource = datasource
        self._database = datasource.DataBase
        self._close = close
        self._email = email
        self._user = IspdbUser()
        self._servers = IspdbServer()
        self._metadata = ''
        self._count = 0
        self._offline = 0
        self._isnew = False
        self._diposed = False
        self._updated = False
        self._version = 0
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
                           'PageLabel': 'IspdbPage%s.Label1.Label',
                           'PagesLabel': 'IspdbPages.Label1.Label',
                           'Progress': 'IspdbPage2.Label2.Label.%s',
                           'Security': 'IspdbPages.Label10.Label.%s',
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
    def getServerConfig(self, *args):
        self._version += 1
        url = self._configuration.getByName('IspDBUrl')
        self._datasource.getServerConfig(self.Email, url, *args)

    def setServerConfig(self, user, smtp, imap, offline):
        self._user.setUser(user)
        self._servers.setServers(smtp, imap, self._user)
        self._offline = offline

# IspdbModel getter methods called by WizardPage3
    def refreshView(self, version):
        return version < self._version

    def getVersion(self):
        return self._version

    def getConfig(self, service):
        config = KeyMap()
        config.setValue('First', self._servers.isFirst(service))
        config.setValue('Last', self._servers.isLast(service))
        config.setValue('Page', self._servers.getServerPage(service))
        config.setValue('Default', self._servers.isDefaultPage(service, self._user))
        config.update(self._servers.getCurrentServer(service))
        config.setValue('Login', self._getLoginName(service))
        config.setValue('Password', self._getPassword(service))
        return config

    def getAuthentication(self, service):
        return self._servers.getServerAuthentication(service)

    def isConnectionValid(self, host, port):
        return self._isHostValid(host) and self._isPortValid(port)

    def isStringValid(self, value):
        if validators.length(value, min=1):
            return True
        return False

    def previousServerPage(self, service, server):
        self._servers.getCurrentServer(service).update(server)
        self._servers.decrementIndex(service)

    def nextServerPage(self, service, server):
        self._servers.getCurrentServer(service).update(server)
        self._servers.incrementIndex(service)

    def updateConfiguration(self, service, server, user):
        print("PageModel.updateConfiguration() %s - %s" % (server, user))
        self._servers.updateCurrentServer(service, server)
        print("PageModel.updateConfiguration() server:\n%s\n%s" % (server.toJson(), self._servers._metadata[service]))
        self._user.updateUser(user)
        print("PageModel.updateConfiguration() user:\n%s\n%s" % (user.toJson(), self._user._metadata))

    def saveConfiguration(self, service):
        server = self._getServer()
        if self._isnew or server.toJson() != self._getServerMetaData():
            provider = self._getDomain()
            host, port = self._getServerKeys()
            self._saveServer(provider, host, port, server)
            print("PageModel.saveConfiguration() server:\n%s\n%s" % (server.toJson(), self._metadata['Servers'][self._index]))
        if self._user.isUpdated():
            print("PageModel.saveConfiguration() user:\n%s\n%s" % (self._user.getUser().toJson(), self._user._metadata))
            self._database.mergeUser(self.Email, self._user.getUser())

    def getSecurity(self, i, j):
        level = self._levels.get(i).get(j)
        message = self.getSecurityMessage(level)
        return message, level

    def _saveServer(self, provider, host, port, server):
        if self._isnew:
            self._database.mergeProvider(provider)
            self._database.mergeServer(provider, server)
        else:
            self._database.updateServer(host, port, server)

    def _getSmtpConfig1(self, email, url, progress, updateModel):
        progress(5)
        url = getUrl(self._ctx, url)
        progress(10)
        mode = getConnectionMode(self._ctx, url.Server)
        progress(20)
        self.waitForDataBase()
        progress(40)
        user, servers = self._database.getSmtpConfig(email)
        if len(servers) > 0:
            progress(100, 1)
        elif mode == OFFLINE:
            progress(100, 2)
        else:
            progress(60)
            service = 'com.gmail.prrvchr.extensions.OAuth2OOo.OAuth2Service'
            request = createService(self._ctx, service)
            response = self._getIspdbConfig(request, url.Complete, user.getValue('Domain'))
            if response.IsPresent:
                progress(80)
                servers = self._database.setSmtpConfig(response.Value)
                progress(100, 3)
            else:
                progress(100, 4)
        updateModel(user, servers, mode)

    def _getIspdbConfig(self, request, url, domain):
        parameter = uno.createUnoStruct('com.sun.star.auth.RestRequestParameter')
        parameter.Method = 'GET'
        parameter.Url = '%s%s' % (url, domain)
        parameter.NoAuth = True
        #parameter.NoVerify = True
        response = request.getRequest(parameter, DataParser()).execute()
        return response

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
    def _getServer(self, service):
        return self._servers.getCurrentServer(service)

    def _getDomain(self):
        return self._user.getDomain()

# IspdbModel private getter methods called by WizardPage3
    def _getLoginName(self, service):
        login = self._user.getLoginName(service)
        return login if login else self._getLoginFromEmail(service)

    def _getLoginFromEmail(self, service):
        mode = self._servers.getLoginMode(service)
        return self.Email.partition('@')[mode] if mode != 1 else self.Email

    def _getPassword(self, service):
        return self._user.getPassword(service)

    def _isHostValid(self, host):
        if validators.domain(host):
            return True
        return False

    def _isPortValid(self, port):
        if validators.between(port, min=1, max=1023):
            return True
        return False

    def _getServerMetaData(self):
        return self._metadata['Servers'][self._index]

    def _getServerKeys(self):
        server = json.loads(self._getServerMetaData())
        return server['Server'], server['Port']

# IspdbModel private getter methods called by WizardPage4
    def _getConnectionContext(self, service):
        server = self._servers.getServerHost(service)
        port = self._server.getServerPort(service)
        connection = self._getConnectionMap(service)
        authentication = self._getAuthenticationMap(service)
        data = {'ServerName': server,
                'Port': port,
                'ConnectionType': connection,
                'AuthenticationType': authentication,
                'Timeout': self.Timeout}
        return CurrentContext(data)

    def _getAuthenticator(self, service):
        data = {'LoginName': self._user.getLoginName(service),
                'Password': self._user.getPassword(service)}
        return Authenticator(data)

    def _getConnectionMap(self, service):
        index = self._server.getServerConnection(service)
        return self._connections.get(index)

    def _getAuthenticationMap(self, service):
        index = self._server.getServerAuthentication(service)
        return self._authentications.get(index)

# IspdbModel StringResource methods
    def getPageStep(self, pageid):
        resource = self._resources.get('Step') % pageid
        return self._resolver.resolveString(resource)

    def getPageTitle(self, pageid):
        resource = self._resources.get('Title') % (pageid, self._offline)
        return self._resolver.resolveString(resource)

    def getPageLabel(self, pageid):
        resource = self._resources.get('PageLabel') % pageid
        return self._resolver.resolveString(resource)

    def getPagesLabel(self, service):
        resource = self._resources.get('PagesLabel')
        label = self._resolver.resolveString(resource)
        return label % (service, self.Email)

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
