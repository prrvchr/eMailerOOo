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

from com.sun.star.frame import XDispatchProvider

from com.sun.star.lang import XInitialization
from com.sun.star.lang import XServiceInfo

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from emailer import SmtpDispatch

from emailer import createMessageBox
from emailer import getExtensionVersion
from emailer import getOAuth2Version
from emailer import getStringResource

from emailer import g_identifier
from emailer import g_extension
from emailer import g_resource
from emailer import g_basename

from packaging import version
import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = '%s.SmtpDispatcher' % g_identifier


class SmtpDispatcher(unohelper.Base,
                     XDispatchProvider,
                     XInitialization,
                     XServiceInfo):
    def __init__(self, ctx):
        self._ctx = ctx
        self._frame = None

# XInitialization
    def initialize(self, args):
        if len(args) > 0:
            self._frame = args[0]

# XDispatchProvider
    def queryDispatch(self, url, frame, flags):
        dispatch = None
        if url.Path in ('ispdb', 'spooler', 'mailer', 'merger'):
            parent = self._frame.getContainerWindow()
            oauth2 = getOAuth2Version(self._ctx)
            driver = getExtensionVersion(self._ctx, 'io.github.prrvchr.jdbcDriverOOo')
            if oauth2 is None:
                self._showMsgBox(parent, 500, 'OAuth2OOo', g_extension)
            elif not self._checkVersion(oauth2, '1.1.1'):
                self._showMsgBox(parent, 502, oauth2, 'OAuth2OOo', '1.1.1')
            elif driver is None:
                self._showMsgBox(parent, 500, 'jdbcDriverOOo', g_extension)
            elif not self._checkVersion(driver, '1.0.5'):
                self._showMsgBox(parent, 502, driver, 'jdbcDriverOOo', '1.0.5')
            else:
                dispatch = SmtpDispatch(self._ctx, parent)
        return dispatch

    def queryDispatches(self, requests):
        dispatches = []
        for request in requests:
            dispatch = self.queryDispatch(request.FeatureURL, request.FrameName, request.SearchFlags)
            dispatches.append(dispatch)
        return tuple(dispatches)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    # Private methods
    def _showMsgBox(self, parent, code, *args):
        resource = getStringResource(self._ctx, g_identifier, g_resource, g_basename)
        title = resource.resolveString(code +1) % g_extension
        message = resource.resolveString(code +2) % args
        msgbox = createMessageBox(parent, message, title, 'error', 1)
        msgbox.execute()
        msgbox.dispose()

    def _checkVersion(self, ver, minimum):
        return version.parse(ver) >= version.parse(minimum)


g_ImplementationHelper.addImplementation(SmtpDispatcher,                            # UNO object class
                                         g_ImplementationName,                      # Implementation name
                                        (g_ImplementationName,))                    # List of implemented services
