#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.awt import XContainerWindowEventHandler

from .pagecontroller import PageController

import traceback


class WizardHandler(unohelper.Base,
                    XContainerWindowEventHandler):
    def __init__(self, ctx, wizard, model):
        self.ctx = ctx
        self._controller = PageController(self.ctx, wizard, model)
        print("WizardHandler.__init__()")

    def getController(self):
        return self._controller

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        try:
            handled = False
            if method == 'TextChange':
                self._controller.updateTravelUI(window, event.Source)
                handled = True
            elif method == 'ChangeAuthentication':
                self._controller.changeAuthentication(window, event.Source)
                handled = True
            elif method == 'Previous':
                self._controller.previousServerPage(window)
                handled = True
            elif method == 'Next':
                self._controller.nextServerPage(window)
                handled = True
            elif method == 'SmtpConnect':
                self._controller.smtpConnect(window, event.Source)
                handled = True
            return handled
        except Exception as e:
            print("WizardHandler.callHandlerMethod() ERROR: %s - %s" % (e, traceback.print_exc()))
    def getSupportedMethodNames(self):
        return ('TextChange', 'ChangeAuthentication', 'Previous', 'Next', 'SmtpConnect')
