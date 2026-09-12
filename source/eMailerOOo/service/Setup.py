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

import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.task import XAsyncJob

from emailer import SetupManager

from emailer import createMessageBox
from emailer import getStringResource

from emailer import g_identifier


import socket
import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'io.github.prrvchr.eMailerOOo.Setup'
g_ServiceNames = ('io.github.prrvchr.eMailerOOo.Setup',
                  'com.sun.star.task.Job')


class Setup(unohelper.Base,
            XServiceInfo,
            XAsyncJob):
    def __init__(self, ctx):
        self._ctx = ctx
        self._job = "eMailerOOoSetup"
        self._name = 'SetupWindow'
        self._code = 600
        self._resources = {'Title': 'Setup.ErrorBox.Title',
                           'Message': 'Setup.ErrorBox.Message'}

    # XAsyncJob
    def executeAsync(self, arguments, listener):
        try:
            if self._checkInternet():
                SetupManager(self._ctx, self._job, self._name, self._code)
            else:
                self._showMessageBox()
        except Exception as e:
            # FIXME: It is essential to notify LibreOffice of
            # FIXME: the Job's completion so as not to block its loading.
            pass
        finally:
            if listener is not None:
                listener.jobFinished(self, None)
        return None

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    # Show MessageBox Error
    def _showMessageBox(self):
        resolver = getStringResource(self._ctx, g_identifier, 'dialogs', 'MessageBox')
        title = resolver.resolveString(self._resources.get('Title'))
        message = resolver.resolveString(self._resources.get('Message'))
        dialog = createMessageBox(self._ctx, title, message)
        dialog.execute()
        dialog.dispose()

    def _checkInternet(self, host="8.8.8.8", port=53, timeout=3):
        try:
            socket.setdefaulttimeout(timeout)
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.connect((host, port))
            return True
        except (OSError, socket.timeout):
            return False


g_ImplementationHelper.addImplementation(Setup,
                                         g_ImplementationName,
                                         g_ServiceNames)

