#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-25 https://prrvchr.github.io                                  ║
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

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail import XMailService
from com.sun.star.lang import EventObject

g_message = 'MailServiceProvider'

import sys
import poplib
import traceback


class Pop3Service(unohelper.Base,
                  XMailService):
    def __init__(self, ctx):
        self._ctx = ctx
        self._listeners = []
        self._supportedconnection = ('Insecure', 'SSL', 'TLS')
        self._supportedauthentication = ('None', 'Login')
        self._server = None
        self._context = None

    def addConnectionListener(self, listener):
        self.listeners.append(listener)

    def removeConnectionListener(self, listener):
        if listener in self._listeners:
            self.listeners.remove(listener)

    def getSupportedConnectionTypes(self):
        return self._supportedconnection

    def getSupportedAuthenticationTypes(self):
        return self._supportedauthentication

    def getCurrentConnectionContext(self):
        return self._context

    def connect(self, context, authenticator):
        self._context = context
        server = context.getValueByName('ServerName')
        port = context.getValueByName('Port')
        timeout = context.getValueByName('Timeout')
        connection = context.getValueByName('ConnectionType').upper()
        authentication = context.getValueByName('AuthenticationType').upper()
        if connection == 'SSL':
            self._server = poplib.POP3_SSL(host=server, port=port, timeout=timeout)
        else:
            self._server = poplib.POP3(host=server, port=port, timeout=timeout)
        if connection == 'TLS':
            self._server.stls()
        if authentication == 'LOGIN':
            user = authenticator.getUserName()
            password = authenticator.getPassword()
            if user != '':
                if sys.version < '3': # fdo#59249 i#105669 Python 2 needs "ascii"
                    user = user.encode('ascii')
                    if password != '':
                        password = password.encode('ascii')
                self._server.user(user)
                self._server.pass_(password)
        notify = EventObject(self)
        for listener in self._listeners:
            listener.connected(notify)

    def disconnect(self):
        if self.isConnected():
            self._server.quit()
            self._server = None
            self._context = None
            notify = EventObject(self)
            for listener in self._listeners:
                listener.disconnected(notify)

    def isConnected(self):
        return self._server is not None
