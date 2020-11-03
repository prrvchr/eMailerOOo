#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.awt import XContainerWindowEventHandler

from .pagemanager import PageManager

import traceback


class PageHandler(unohelper.Base,
                  XContainerWindowEventHandler):
    def __init__(self, ctx, wizard, model):
        self.ctx = ctx
        self._enabled = True
        self._manager = PageManager(self.ctx, wizard, model)
        print("PageHandler.__init__()")

    def disable(self):
        self._enabled = False

    def enable(self):
        self._enabled = True

    def getManager(self):
        return self._manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        try:
            handled = False
            if method == 'TextChange':
                if self._enabled:
                    self._manager.updateTravelUI()
                handled = True
            elif method == 'ChangeConnection':
                self._manager.changeConnection(window, event.Source)
                handled = True
            elif method == 'ChangeAuthentication':
                self._manager.changeAuthentication(window, event.Source)
                handled = True
            elif method == 'Previous':
                self._manager.previousServerPage(window)
                handled = True
            elif method == 'Next':
                self._manager.nextServerPage(window)
                handled = True
            elif method == 'SmtpConnect':
                self._manager.smtpConnect(window)
                handled = True
            return handled
        except Exception as e:
            print("PageHandler.callHandlerMethod() ERROR: %s - %s" % (e, traceback.print_exc()))
    def getSupportedMethodNames(self):
        return ('TextChange', 'ChangeConnection', 'ChangeAuthentication',
                'Previous', 'Next', 'SmtpConnect')
