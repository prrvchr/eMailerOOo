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
                self._manager.updateTravelUI()
            handled = True
        elif method == 'ChangeConnection':
            self._manager.changeConnection(event.Source)
            handled = True
        elif method == 'ChangeAuthentication':
            self._manager.changeAuthentication(event.Source)
            handled = True
        elif method == 'Previous':
            self._manager.previousServerPage()
            handled = True
        elif method == 'Next':
            self._manager.nextServerPage()
            handled = True
        elif method == 'SendMail':
            self._manager.sendMail()
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('TextChange', 'ChangeConnection', 'ChangeAuthentication',
                'Previous', 'Next', 'SendMail')


class DialogHandler(unohelper.Base,
                    XDialogEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'TextChange':
            self._manager.updateDialog()
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('TextChange', )
