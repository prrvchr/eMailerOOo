#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.awt import XDialogEventHandler

import traceback


class PageHandler(unohelper.Base,
                  XContainerWindowEventHandler):
    def __init__(self, manager):
        self._enabled = True
        self._manager = manager

    @property
    def Manager(self):
        return self._manager

    def disable(self):
        self._enabled = False

    def enable(self):
        self._enabled = True

    def getManager(self):
        return self.Manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        handled = False
        if method == 'TextChange':
            if self._enabled:
                self.Manager.updateTravelUI()
            handled = True
        elif method == 'ChangeConnection':
            self.Manager.changeConnection(event.Source)
            handled = True
        elif method == 'ChangeAuthentication':
            self.Manager.changeAuthentication(event.Source)
            handled = True
        elif method == 'Previous':
            self.Manager.previousServerPage()
            handled = True
        elif method == 'Next':
            self.Manager.nextServerPage()
            handled = True
        elif method == 'ShowSmtpConnect':
            self.Manager.showSmtpConnect()
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('TextChange', 'ChangeConnection', 'ChangeAuthentication',
                'Previous', 'Next', 'ShowSmtpConnect')


class DialogHandler(unohelper.Base,
                    XDialogEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'SmtpConnect':
            self._manager.smtpConnect()
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('SmtpConnect', )
