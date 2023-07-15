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
    def __init__(self):
        services = {SMTP.value: 25, IMAP.value: 143}
        self._servers = {service: [] for service in services}
        self._metadata = {service: () for service in services}
        self._index = {service: -1 for service in services}
        self._default = {service: port for service, port in services.items()}

    def appendServer(self, service, server):
        self._servers[service].append(server)

    def appendServers(self, service, servers):
        self._servers[service] = servers

    def hasServers(self, services):
        for service in services:
            if self._getServerCount(service) == 0:
                return False
        return True

    def setDefault(self, services, user):
        for service in services:
            servers = self._servers[service]
            self._setDefault(service, servers, user)
        return self

    def _setDefault(self, service, servers, user):
        if len(servers) > 0:
            self._metadata[service] = tuple(json.dumps(server, sort_keys=True) for server in servers)
        else:
            self._metadata[service] = ()
            servers = self._getDefaultServers(service, user)
        self._servers[service] = servers
        self._index[service] = self._getServerIndex(service, user, 0)

    def _isNew(self, service):
        return len(self._metadata[service]) == 0

    def getCurrentServer(self, service):
        return self._servers[service][self._getIndex(service)]

    def updateCurrentServer(self, service, server):
        self.getCurrentServer(service).update(server)

    def _getDefaultServers(self, service, user):
        host = '%s.%s' % (service.lower(), user.getDomain())
        server = {}
        server['Service'] = service
        server['Server'] = host
        server['Port'] = self._default[service]
        server['Connection'] = 0
        server['Authentication'] = 0
        server['LoginMode'] = 1
        return (server, )

    def _getIndex(self, service):
        return self._index[service]

    def incrementIndex(self, service):
        self._index[service] += 1

    def decrementIndex(self, service):
        self._index[service] -= 1

    def _getServerIndex(self, service, user, default=-1):
        server = user.getServer(service)
        port = user.getPort(service)
        servers = self._servers[service]
        if port != 0:
            for s in servers:
                if s.get('Server') == server and s.get('Port') == port:
                    default = servers.index(s)
                    break;
        return default

    def isFirst(self, service):
        return self._getIndex(service) == 0

    def isLast(self, service):
        count = self._getServerCount(service)
        return self._getIndex(service) +1 >= count

    def getServerPage(self, service):
        if self._isNew(service):
            page = '1/0'
        else:
            count = self._getServerCount(service)
            page = '%s/%s' % (self._getIndex(service) +1, count)
        return page

    def isDefaultPage(self, service, user):
        return self._getIndex(service) == self._getServerIndex(service, user)

    def getLoginMode(self, service):
        return self.getCurrentServer(service).get('LoginMode')

    def getServerHost(self, service):
        return self.getCurrentServer(service).get('Server')
        
    def getServerPort(self, service):
        return self.getCurrentServer(service).get('Port')

    def getServerConnection(self, service):
        return self.getCurrentServer(service).get('Connection')
        
    def getServerAuthentication(self, service):
        return self.getCurrentServer(service).get('Authentication')

    def getConfig(self, service, timeout, connections, authentications):
        config = {}
        config[service + 'ServerName'] = self.getServerHost(service)
        config[service + 'Port'] = self.getServerPort(service)
        config[service + 'ConnectionType'] = connections.get(self.getServerConnection(service))
        config[service + 'AuthenticationType'] = authentications.get(self.getServerAuthentication(service))
        config[service + 'Timeout'] = timeout
        return config

    def saveServer(self, datasource, service, provider):
        new = self._isNew(service)
        server = self.getCurrentServer(service)
        metadata = self._getServerMetaData(service)
        if new or json.dumps(server, sort_keys=True) != metadata:
            host, port = self._getServerKeys(metadata)
            datasource.saveServer(new, provider, server, host, port)

    def _getServerMetaData(self, service):
        return self._metadata[service][self._getIndex(service)]

    def _getServerKeys(self, metadata):
        server = json.loads(metadata)
        return server['Server'], server['Port']

    def _getServerCount(self, service):
        return len(self._servers[service])
