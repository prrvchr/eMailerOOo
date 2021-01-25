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

import unohelper

from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt import XContainerWindowEventHandler

import traceback


class DialogHandler(unohelper.Base,
                    XDialogEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'Send':
            self._manager.sendDocument()
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('Send', )

class WindowHandler(unohelper.Base,
                    XContainerWindowEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'AddSender':
            self._manager.addSender()
            handled = True
        elif method == 'RemoveSender':
            self._manager.removeSender()
            handled = True
        elif method == 'ChangeRecipient':
            self._manager.changeRecipient()
            handled = True
        elif method == 'AddRecipient':
            self._manager.addRecipient()
            handled = True
        elif method == 'RemoveRecipient':
            self._manager.removeRecipient()
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('AddSender', 'RemoveSender', 'ChangeRecipient', 'AddRecipient', 'RemoveRecipient')
