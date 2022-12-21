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
from com.sun.star.frame import XDispatchResultListener
from com.sun.star.mail import XIspDBService

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpmailer.unotool import getUrl
from smtpmailer.unotool import createService
from smtpmailer.unotool import getNamedValueSet

from smtpmailer.logger import logMessage
from smtpmailer.logger import getMessage

from smtpmailer import DataSource
from smtpmailer import g_identifier

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = '%s.IspDBService' % g_identifier


class IspDBService(unohelper.Base,
                   XServiceInfo,
                   XIspDBService,
                   XDispatchResultListener):
    def __init__(self, ctx, email=None):
        self.ctx = ctx
        self._context = None
        self._authenticator = None
        self._datasource = DataSource(self.ctx)
        if email is not None:
            self.initialize(email)
        logMessage(self.ctx, INFO, "Loading ... Done", 'IspDBService', '__init__()')

    # XIspDBService
    def initialize(self, email):
        print("IspDBService.initialize() 1")
        url = getUrl(self.ctx, 'ispdb://')
        desktop = createService(self.ctx, 'com.sun.star.frame.Desktop')
        dispatcher = desktop.getCurrentFrame().queryDispatch(url, '', 0)
        print("IspDBService.initialize() 2")
        if dispatcher is not None:
            args = getNamedValueSet({'Email': email})
            dispatcher.dispatchWithNotification(url, args, self)
            print("IspDBService.initialize() 3")

    def getConnectionContext(self):
        self._datasource.waitForConfig()
        return self._context
    def getAuthenticator(self):
        self._datasource.waitForConfig()
        return self._authenticator

    # XDispatchResultListener
    def dispatchFinished(self, result):
        print("IspDBService.dispatchFinished() %s" % (result, ))
    def disposing(self, source):
        pass

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)


g_ImplementationHelper.addImplementation(IspDBService,                              # UNO object class
                                         g_ImplementationName,                      # Implementation name
                                        (g_ImplementationName,))                    # List of implemented services
