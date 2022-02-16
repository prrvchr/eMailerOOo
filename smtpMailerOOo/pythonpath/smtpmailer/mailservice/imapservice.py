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

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.lang import EventObject
from com.sun.star.mail import XImapService

from smtpmailer import getMessage
from smtpmailer import isDebugMode
from smtpmailer import logMessage

g_message = 'MailServiceProvider'

import traceback


class ImapService(unohelper.Base,
                  XImapService):
    def __init__(self, ctx):
        if isDebugMode():
            msg = getMessage(ctx, g_message, 311)
            logMessage(ctx, INFO, msg, 'ImapService', '__init__()')
        self._ctx = ctx
        self._listeners = []
        self._supportedconnection = ('Insecure', 'Ssl', 'Tls')
        self._supportedauthentication = ('None', 'Login', 'OAuth2')
        self._server = None
        self._context = None
        self._notify = EventObject(self)
        if isDebugMode():
            msg = getMessage(ctx, g_message, 312)
            logMessage(ctx, INFO, msg, 'ImapService', '__init__()')

# XMailService2 interface implementation
    def addConnectionListener(self, listener):
        self._listeners.append(listener)

    def removeConnectionListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

    def getSupportedConnectionTypes(self):
        return self._supportedconnection

    def getSupportedAuthenticationTypes(self):
        return self._supportedauthentication

    def getCurrentConnectionContext(self):
        return self._context

# Interface not implemented
    def connect(self, context, authenticator):
        raise NotImplementedError

    def disconnect(self):
        raise NotImplementedError

    def isConnected(self):
        raise NotImplementedError

    def getSentFolder(self):
        raise NotImplementedError

    def hasFolder(self, folder):
        raise NotImplementedError

    def uploadMessage(self, folder, message):
        raise NotImplementedError
