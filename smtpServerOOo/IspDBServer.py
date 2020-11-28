#!
# -*- coding: utf_8 -*-

'''
    Copyright (c) 2020 https://prrvchr.github.io

    Permission is hereby granted, free of charge, to any person obtaining
    a copy of this software and associated documentation files (the "Software"),
    to deal in the Software without restriction, including without limitation
    the rights to use, copy, modify, merge, publish, distribute, sublicense,
    and/or sell copies of the Software, and to permit persons to whom the Software
    is furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
    EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
    OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
    CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
    TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
    OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'''

import uno
import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.lang import XInitialization
from com.sun.star.frame import XDispatchProvider

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpserver import PageModel
from smtpserver import IspdbDispatch

from smtpserver import logMessage
from smtpserver import getMessage

from smtpserver import g_identifier

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = '%s.IspDBServer' % g_identifier


class IspDBServer(unohelper.Base,
                  XServiceInfo,
                  XInitialization,
                  XDispatchProvider):
    def __init__(self, ctx):
        self.ctx = ctx
        self._frame = None
        self._model = PageModel(self.ctx)
        logMessage(self.ctx, INFO, "Loading ... Done", 'IspDBServer', '__init__()')

    # XInitialization
    def initialize(self, args):
        if len(args) > 0:
            print("IspDBServer.initialize()")
            self._frame = args[0]

    # XDispatchProvider
    def queryDispatch(self, url, frame, flags):
        if url.Protocol != 'ispdb:':
            print("IspDBServer.queryDispatch() 1 %s" % url.Protocol)
            return None
        print("IspDBServer.queryDispatch()2")
        dispatch = IspdbDispatch(self.ctx, self._model, self._frame)
        return dispatch
    def queryDispatches(self, requests):
        dispatches = []
        for r in requests:
            dispatches.append(self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags))
        return tuple(dispatches)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)


g_ImplementationHelper.addImplementation(IspDBServer,                               # UNO object class
                                         g_ImplementationName,                      # Implementation name
                                        (g_ImplementationName,))                    # List of implemented services
