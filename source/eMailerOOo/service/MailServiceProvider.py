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

from com.sun.star.lang import XServiceInfo
from com.sun.star.mail import XMailServiceProvider2

from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.mail.MailServiceType import POP3
from com.sun.star.mail.MailServiceType import IMAP

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.mail import NoMailServiceProviderException

from emailer import SmtpApiService
from emailer import SmtpBaseService
from emailer import Pop3Service
from emailer import ImapApiService
from emailer import ImapBaseService

from emailer import getLogger

from emailer import g_mailservicelog

g_message = 'MailServiceProvider'

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplName = 'com.sun.star.mail.MailServiceProvider2'


class MailServiceProvider(unohelper.Base,
                          XMailServiceProvider2,
                          XServiceInfo):
    def __init__(self, ctx):
        self._logger = getLogger(ctx, g_mailservicelog)
        self._logger.logprb(INFO, 'MailServiceProvider', '__init__()', 111)
        self._ctx = ctx
        self._logger.logprb(INFO, 'MailServiceProvider', '__init__()', 112)

    def create(self, mailtype, host):
        self._logger.logprb(INFO, 'MailServiceProvider', 'create()', 121, mailtype.value)
        if mailtype == SMTP:
            if host.endswith('gmail.com'):
                service = SmtpApiService(self._ctx)
            else:
                service = SmtpBaseService(self._ctx)
        elif mailtype == POP3:
            service = Pop3Service(self._ctx)
        elif mailtype == IMAP:
            if host.endswith('gmail.com'):
                service = ImapApiService(self._ctx)
            else:
                service = ImapBaseService(self._ctx)
        else:
            e = self._getNoMailServiceProviderException(123, mailtype)
            self._logger.logp(SEVERE, 'MailServiceProvider', 'create()', e.Message)
            raise e
        self._logger.logprb(INFO, 'MailServiceProvider', 'create()', 122, mailtype.value)
        return service

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplName, service)
    def getImplementationName(self):
        return g_ImplName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplName)

    def _getNoMailServiceProviderException(self, code, *args):
        e = NoMailServiceProviderException()
        e.Message = self._logger.resolveString(code, *args)
        e.Context = self
        return e


g_ImplementationHelper.addImplementation(MailServiceProvider,
                                         g_ImplName,
                                         ('com.sun.star.mail.MailServiceProvider2', ))