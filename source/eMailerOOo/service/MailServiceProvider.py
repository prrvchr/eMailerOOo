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

from com.sun.star.lang import XServiceInfo

from com.sun.star.mail import XMailServiceProvider
from com.sun.star.mail import NoMailServiceProviderException

from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.mail.MailServiceType import POP3
from com.sun.star.mail.MailServiceType import IMAP

from com.sun.star.logging.LogLevel import ALL
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from emailer import ImapService
from emailer import Pop3Service
from emailer import SmtpService

from emailer import getLogger

from emailer import getConfiguration

from emailer import g_identifier
from emailer import g_mailservicelog

from email.utils import parseaddr
import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'com.sun.star.mail.MailServiceProvider'
g_ServiceNames = ('com.sun.star.mail.MailServiceProvider', )


class MailServiceProvider(unohelper.Base,
                          XMailServiceProvider,
                          XServiceInfo):
    def __init__(self, ctx):
        self._ctx = ctx
        mtd = '__init__'
        self._cls = 'MailServiceProvider'
        logger = getLogger(ctx, g_mailservicelog)
        debug = logger.Level == ALL
        if debug:
            logger.logprb(INFO, self._cls, mtd, 101)
        self._providers = None
        self._provider = 'Providers'
        self._logger = logger
        self._debug = debug
        if debug:
            logger.logprb(INFO, self._cls, mtd, 102)

    def create(self, stype):
        mtd = 'create'
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 111, stype.value)
        if stype == SMTP:
            service = SmtpService(self._ctx, self._logger, self._getProviders(), self._debug)
        elif stype == POP3:
            service = Pop3Service(self._ctx)
        elif stype == IMAP:
            service = ImapService(self._ctx, self._logger, self._getProviders(), self._debug)
        else:
            e = self._getNoMailServiceProviderException(112, stype.value)
            if self._debug:
                self._logger.logp(SEVERE, self._cls, mtd, e.Message)
            raise e
        if self._debug:
            self._logger.logprb(INFO, self._cls, mtd, 113, stype.value)
        return service

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _getProviders(self):
        if self._providers is None:
            config = getConfiguration(self._ctx, g_identifier)
            self._providers = config.getByName(self._provider)
        return self._providers

    def _getNoMailServiceProviderException(self, code, *args):
        e = NoMailServiceProviderException()
        e.Message = self._logger.resolveString(code, *args)
        e.Context = self
        return e

g_ImplementationHelper.addImplementation(MailServiceProvider,             # UNO object class
                                         g_ImplementationName,            # Implementation name
                                         g_ServiceNames)                  # List of implemented services
