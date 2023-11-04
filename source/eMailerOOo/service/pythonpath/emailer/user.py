#!
# -*- coding: utf-8 -*-

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

from .mailerlib import CurrentContext
from .mailerlib import Authenticator

from .unotool import getConfiguration

from .configuration import g_identifier

from email.utils import parseaddr
import traceback


class User(unohelper.Base):
    def __init__(self, ctx, sender, new=False):
        self._sender = sender
        name, address = parseaddr(sender)
        self._email = address
        name, sep, domain = address.partition('@')
        self._domain = domain
        self._services = {SMTP.value: 465, IMAP.value: 993}
        self._config = getConfiguration(ctx, g_identifier, True)
        self._timeout = self._config.getByName('ConnectTimeout')
        senders = self._config.getByName('Senders')
        if senders.hasByName(sender):
            user = senders.getByName(sender)
        else:
            senders.insertByName(sender, senders.createInstance())
            user = senders.getByName(sender)
            user.replaceByName('UseReplyTo', False)
            user.replaceByName('UseIMAP', True)
            user.replaceByName('ReplyToAddress', '')
            servers = user.getByName('Servers')
            self._setDefaultServer(servers, SMTP.value)
            self._setDefaultServer(servers, IMAP.value)
            new = True
        self._new = new
        self._user = user
        self._properties = ('ServerName', 'Port', 'ConnectionType',
                            'AuthenticationType', 'UserName', 'Password')
        self._indexes = {'ConnectionType': ('Insecure', 'SSL', 'TLS'),
                         'AuthenticationType': ('None', 'Login', 'OAuth2')}
        self._sep = '/'

    @property
    def _servers(self):
        return self._user.getByName('Servers')

    @property
    def Sender(self):
        return self._sender

    @property
    def ReplyToAddress(self):
        if self.useReplyTo():
            return self._user.getByName('ReplyToAddress')
        return self._sender

    def useReplyTo(self):
        return self._user.getByName('UseReplyTo')

    def getReplyToState(self):
        return 1 if self.useReplyTo() else 0

    def enableReplyTo(self, enabled):
        if self.useReplyTo() != enabled:
            self._user.replaceByName('UseReplyTo', enabled)

    def setReplyToAddress(self, replyto):
        self._user.replaceByName('ReplyToAddress', replyto)

    def isNew(self):
        return self._new

    def getServices(self):
        return self._services

    def getImapState(self):
        return 1 if self.useIMAP() else 0

    def enableImap(self, enabled):
        if self.useIMAP() != enabled:
            self._user.replaceByName('UseIMAP', enabled)

    def useIMAP(self):
        return self._user.getByName('UseIMAP')

    def updateConfig(self, service, user):
        for key, value in user.items():
            service, sep, property = key.partition(self._sep)
            if property in self._indexes:
                value = self._indexes[property][value]
            if self._servers.hasByHierarchicalName(key):
                if self._servers.getByHierarchicalName(key) != value:
                    self._servers.getByName(service).replaceByName(property, value)
            else:
                if not self._servers.hasByName(service):
                    self._servers.insertByName(service, self._servers.createInstance())
                self._servers.getByName(service).replaceByName(property, value)

    def saveConfig(self):
        if self._config.hasPendingChanges():
            self._config.commitChanges()

    def getServerName(self, service):
        return self._getPropertyValue(service, 'ServerName', self.getHost(service))

    def getPort(self, service):
        return self._getPropertyValue(service, 'Port', self._services[service])

    def getConnectionType(self, service):
        return self._getPropertyValue(service, 'ConnectionType', 1)

    def getAuthenticationType(self, service):
        return self._getPropertyValue(service, 'AuthenticationType', 1)

    def getUserName(self, service):
        return self._getPropertyValue(service, 'UserName', '')

    def getPassword(self, service):
        return self._getPropertyValue(service, 'Password', '')

    def getUserDomain(self):
        return self._domain

    def getServers(self):
        servers = {}
        for service in  self._servers.ElementNames:
            servers[service] = self.getServer(service)
        return servers

    def getServer(self, service):
        servers = []
        if self._servers.hasByName(service):
            server = {}
            server['ServerName'] = self.getServerName(service)
            server['Port'] = self.getPort(service)
            server['ConnectionType'] = self.getConnectionType(service)
            server['AuthenticationType'] = self.getAuthenticationType(service)
            servers.append(server)
        return servers

    def getHost(self, service):
        return '%s.%s' % (service.lower(), self._domain)

    def getConnectionContext(self, service):
        config = {}
        if self._servers.hasByName(service):
            server = self._servers.getByName(service)
            config['ServerName'] = server.getByName('ServerName')
            config['Port'] = server.getByName('Port')
            config['ConnectionType'] = server.getByName('ConnectionType')
            config['AuthenticationType'] = server.getByName('AuthenticationType')
            config['Timeout'] = self._timeout
        return CurrentContext(config)

    def getAuthenticator(self, service):
        config = {}
        if self._servers.hasByName(service):
            server = self._servers.getByName(service)
            config['Login'] = server.getByName('UserName')
            config['Password'] = server.getByName('Password')
        return Authenticator(config)

    # Private method
    def _getPropertyValue(self, service, property, default):
        key = service + self._sep + property
        if self._servers.hasByHierarchicalName(key):
            default = self._servers.getByHierarchicalName(key)
            if property in self._indexes:
                default = self._indexes[property].index(default)
        return default

    def _setPropertyValue(self, user, server, service):
        for property in self._properties:
            value = server.getByName(property)
            if property in self._indexes:
                value = self._indexes[property].index(value)
            user[service + self._sep + property] = value

    def _setDefaultServer(self, servers, service):
        servers.insertByName(service, servers.createInstance())
        server = servers.getByName(service)
        server.replaceByName('ServerName', self.getHost(service))
        server.replaceByName('Port', self._services[service])
        server.replaceByName('ConnectionType', 'SSL')
        server.replaceByName('AuthenticationType', 'Login')
        server.replaceByName('UserName', '')
        server.replaceByName('Password', '')
