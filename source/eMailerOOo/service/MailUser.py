#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-24 https://prrvchr.github.io                                  ║
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

from emailer import User

from emailer import executeDispatch
from emailer import getConfiguration
from emailer import getPropertyValueSet

from emailer import g_identifier

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = f'{g_identifier}.MailUser'


class MailUser(unohelper.Base,
               XServiceInfo,
               XMailUser):
    # FIXME: We should be able to return None if the user does
    # FIXME: not exist and the IspDB Wizard has been canceled.
    def __new__(cls, ctx, sender=''):
        senders = getConfiguration(ctx, g_identifier).getByName('Senders')
        if not senders.hasByName(sender):
            # FIXME: The Sender name must not be able to be changed (ie: ReadOnly)
            arguments = getPropertyValueSet({'Sender': sender, 'ReadOnly': True})
            executeDispatch(ctx, 'smtp:ispdb', arguments)
            # The IspDB Wizard has been canceled
            if not senders.hasByName(sender):
                return None
        return super(MailUser, cls).__new__(cls)

    def __init__(self, ctx, sender):
        self._user = User(ctx, sender)

# XMailUser
    def supportIMAP(self):
        return self._user.useIMAP()

    def useReplyTo(self):
        return self._user.useReplyTo()

    def getReplyToAddress(self):
        return self._user.ReplyToAddress

    def getAuthenticator(self, stype):
        return self._user.getAuthenticator(stype.value)

    def getConnectionContext(self, stype):
        return self._user.getConnectionContext(stype.value)


    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)


g_ImplementationHelper.addImplementation(MailUser,
                                         g_ImplementationName,
                                        ('com.sun.star.mail.MailUser',))

