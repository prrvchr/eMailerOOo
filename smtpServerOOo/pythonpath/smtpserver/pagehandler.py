#!
# -*- coding: utf_8 -*-

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
