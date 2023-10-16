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

import json
import traceback


class IspdbServer(unohelper.Base):
    def __init__(self, user, config, imap):
        services = user.getServices(imap)
        if user.isNew():
            if config is None:
                self._servers = self._getDefaultServers(services, user)
            else:
                self._servers = config
        elif config is None:
            self._servers = user.getServers()
        else:
            self._servers = {}
            for service in config:
                servers = user.getServer(service)
                for server in config[service]:
                    if self._isNewConfig(service, server, user):
                        servers.append(server)
                self._servers[service] = servers
        self._index = {service: 0 for service in self._servers}
        self.User = user

    def getCurrentServer(self, service):
        return self._servers[service][self._getIndex(service)]

    def updateCurrentServer(self, service, server):
        self.getCurrentServer(service).update(server)

    def incrementIndex(self, service):
        self._index[service] += 1

    def decrementIndex(self, service):
        self._index[service] -= 1

    def isFirst(self, service):
        return self._getIndex(service) == 0

    def isLast(self, service):
        count = self._getServerCount(service)
        return self._getIndex(service) +1 >= count

    def getServerPage(self, service):
        count = self._getServerCount(service)
        page = '%s/%s' % (self._getIndex(service) +1, count)
        return page

    def isDefaultPage(self, service):
        return self._getIndex(service) == self._getServerIndex(service)

    def getLoginMode(self, service):
        return self.getCurrentServer(service).get('LoginMode')

    def getServerHost(self, service):
        return self.getCurrentServer(service).get('ServerName')
        
    def getServerPort(self, service):
        return self.getCurrentServer(service).get('Port')

    def getServerConnection(self, service):
        return self.getCurrentServer(service).get('ConnectionType')
        
    def getServerAuthentication(self, service):
        return self.getCurrentServer(service).get('AuthenticationType')

    def getConfig(self, service, timeout, connections, authentications):
        config = {}
        config[service + 'ServerName'] = self.getServerHost(service)
        config[service + 'Port'] = self.getServerPort(service)
        config[service + 'ConnectionType'] = connections.get(self.getServerConnection(service))
        config[service + 'AuthenticationType'] = authentications.get(self.getServerAuthentication(service))
        config[service + 'Timeout'] = timeout
        return config

    def _getDefaultServers(self, services, user):
        servers = {}
        for service, port in services.items():
            servers[service] = (self._getDefaultServer(service, user, port), )
        return servers

    def _getDefaultServer(self, service, user, port):
        server = {}
        server['Service'] = service
        server['ServerName'] = user.getHost(service)
        server['Port'] = user.getPort(service)
        server['ConnectionType'] = 1
        server['AuthenticationType'] = 1
        server['LoginMode'] = 1
        return server

    def _isNewConfig(self, service, server, user):
        return any((server.get('ServerName') != user.getServerName(service),
                    server.get('Port') != user.getPort(service)))

    def _getIndex(self, service):
        return self._index[service]

    def _getServerIndex(self, service, default=-1):
        host = self.User.getServerName(service)
        port = self.User.getPort(service)
        servers = self._servers[service]
        for server in servers:
            if server.get('ServerName') == host and server.get('Port') == port:
                default = servers.index(server)
                break;
        return default

    def _getServerCount(self, service):
        return len(self._servers[service])
