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

import uno
import unohelper

from com.sun.star.mail import XMailUser

from com.sun.star.lang import XServiceInfo

from emailer import DispatchListener
from emailer import User

from emailer import executeDispatch
from emailer import getPropertyValueSet

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'org.openoffice.pyuno.MailUser'


class MailUser(unohelper.Base,
               XServiceInfo,
               XMailUser):
    def __init__(self, ctx, sender=''):
        self._ctx = ctx
        self._sender = sender
        user = User(ctx, sender)
        if user.isNew():
            self._user = None
        else:
            self._user = user
        print("MailUser.__init__() Sender: %s" % sender)

# XMailUser
    def supportIMAP(self):
        if self._user is None:
            self._initUser()
        return self._user.useIMAP()

    def useReplyTo(self):
        if self._user is None:
            self._initUser()
        return self._user.useReplyTo()

    def getReplyToAddress(self):
        if self._user is None:
            self._initUser()
        return self._user.ReplyToAddress

    def getAuthenticator(self, stype):
        if self._user is None:
            self._initUser()
        print("MailUser.getAuthenticator() MailServerType: %s" % stype.value)
        return self._user.getAuthenticator(stype.value)

    def getConnectionContext(self, stype):
        if self._user is None:
            self._initUser()
        print("MailUser.getConnectionContext() MailServerType: %s" % stype.value)
        return self._user.getConnectionContext(stype.value)

# Internal method
    def setUser(self, sender):
        self._user = User(self._ctx, sender)

    def _initUser(self):
        listener = DispatchListener(self.setUser)
        arguments = getPropertyValueSet({'Sender': self._sender})
        executeDispatch(self._ctx, 'smtp:ispdb', arguments, listener)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)


g_ImplementationHelper.addImplementation(MailUser,                                                 # UNO object class
                                         g_ImplementationName,                                     # Implementation name
                                        ('com.sun.star.mail.MailUser',))                           # List of implemented services
