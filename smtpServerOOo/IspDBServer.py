#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.lang import XInitialization
from com.sun.star.frame import XDispatchProvider

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpserver import WizardModel
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
        self._model = WizardModel(self.ctx)
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
