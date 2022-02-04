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

import validators
import json
import traceback


class IspdbServer(unohelper.Base):
    def __init__(self):
        smtp = SMTP.value
        imap = IMAP.value
        self._servers = {}
        self._index = {smtp: -1, imap: -1}
        self._metadata = {}
        self._isnew = {}
        self._default = {smtp: ('smtp.', 25), imap: ('imap.', 143)}

    def getServers(self, service):
        servers = []
        if service in self._servers:
            servers = self._servers[service]
        return servers

    def getServerCount(self, service):
        return len(self.getServers(service))

    def setServers(self, smtp, imap, user):
        self._setServers(SMTP.value, smtp, user)
        self._setServers(IMAP.value, imap, user)

    def _setServers(self, service, servers, user):
        self._isnew[service] = False
        if not len(servers):
            self._isnew[service] = True
            servers = self._getDefaultServers(service)
        self._metadata[service] = tuple(server.toJson() for server in servers)
        self._servers[service] = servers
        self._index[service] = self._getServerIndex(service, user, 0)

    def getCurrentServer(self, service):
        return self._servers[service][self.getIndex(service)]

    def updateCurrentServer(self, service, server):
        self._servers[service][self.getIndex(service)].update(server)

    def _getDefaultServer(self, service):
        server, port = self._default[service]
        server = KeyMap()
        server.setValue('Service', service)
        server.setValue('Server', server)
        server.setValue('Port', port)
        server.setValue('Connection', 0)
        server.setValue('Authentication', 0)
        server.setValue('LoginMode', 1)
        return (server, )

    def setIndex(self, service, index):
        self._index[service] = index

    def getIndex(self, service):
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
                if s.getValue('Server') == server and s.getValue('Port') == port:
                    default = servers.index(s)
                    break;
        return default

    def isFirst(self, service):
        return self._index[service] == 0

    def isLast(self, service):
        count = len(self._servers[service])
        return self._index[service] +1 >= count

    def getServerPage(self, service):
        if self._isnew[service]:
            page = '1/0'
        else:
            count = len(self._servers[service])
            page = '%s/%s' % (self._index[service] +1, count)
        return page

    def isDefaultPage(self, service, user):
        return self._index[service] == self._getServerIndex(service, user)

    def getLoginMode(self, service):
        return self.getCurrentServer(service).getValue('LoginMode')

    def getServerHost(self, service):
        return self.getCurrentServer(service).getValue('Server')
        
    def getServerPort(self, service):
        return self.getCurrentServer(service).getValue('Port')

    def getServerConnection(self, service):
        return self.getCurrentServer(service).getValue('Connection')
        
    def getServerAuthentication(self, service):
        return self.getCurrentServer(service).getValue('Authentication')
