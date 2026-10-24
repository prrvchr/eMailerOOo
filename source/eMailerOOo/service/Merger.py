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

from com.sun.star.awt.MessageBoxType import WARNINGBOX
from com.sun.star.awt import Point

from com.sun.star.lang import XServiceInfo
from com.sun.star.task import XAsyncJob

from emailer import createMessageBox
from emailer import findFrame
from emailer import getConfiguration
from emailer import getDesktop
from emailer import getJobArguments
from emailer import getStringResource
from emailer import notifyDispatch
from emailer import showMergerWizard

from emailer import g_mergerframe
from emailer import g_identifier
from emailer import g_resource

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'io.github.prrvchr.eMailerOOo.Merger'
g_ServiceNames = ('io.github.prrvchr.eMailerOOo.Merger',
                  'com.sun.star.task.Job')


class Merger(unohelper.Base,
             XServiceInfo,
             XAsyncJob):
    def __init__(self, ctx):
        self._ctx = ctx
        self._point = Point(*getConfiguration(ctx, g_identifier).getByName('MergerPosition'))

    # XAsyncJob
    def executeAsync(self, arguments, listener):
        try:
            kwargs = getJobArguments(arguments)
            frame = findFrame(self._ctx, g_mergerframe)
            if frame:
                frame.getContainerWindow().toFront()
                notifyDispatch(self, kwargs.get('dispatch'))
            else:
                frame = getDesktop(self._ctx).getCurrentFrame()
                if frame:
                    document = frame.getController().getModel()
                    if document.hasLocation():
                        showMergerWizard(self._ctx, document, point=self._point)
                    else:
                        self._showMessageBox(1221, document.Title)
                else:
                    self._showMessageBox(1211)
        except Exception as e:
            # FIXME: It is essential to notify LibreOffice of
            # FIXME: the Job's completion so as not to block its loading.
            print("Merger.executeAsync() ERROR: %s" % traceback.format_exc())
            pass
        finally:
            if listener is not None:
                listener.jobFinished(self, None)
        return None

    def _showMessageBox(self, code, *args):
        resolver = getStringResource(self._ctx, g_identifier, g_resource, 'eMailer')
        title = resolver.resolveString(code)
        message = resolver.resolveString(code + 1).format(args)
        dialog = createMessageBox(self._ctx, title, message, WARNINGBOX)
        dialog.execute()
        dialog.dispose()

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)


g_ImplementationHelper.addImplementation(Merger,
                                         g_ImplementationName,
                                         g_ServiceNames)

