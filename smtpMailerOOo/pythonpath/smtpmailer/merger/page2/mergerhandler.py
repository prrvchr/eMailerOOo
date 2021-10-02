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
from com.sun.star.sdbc import XRowSetListener

import traceback


class Tab1Handler(unohelper.Base,
                  XContainerWindowEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        try:
            handled = False
            enabled = self._manager.isHandlerEnabled()
            if method == 'ChangeAddressTable':
                print("Tab1Handler.callHandlerMethod() ChangeAddressTable *************** %s" % enabled)
                if enabled:
                    control = event.Source
                    table = control.getSelectedItem()
                    self._manager.setAddressTable(table)
                handled = True
            elif method == 'Add':
                self._manager.addItem()
                handled = True
            elif method == 'AddAll':
                self._manager.addAllItem()
                handled = True
            return handled
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)

    def getSupportedMethodNames(self):
        return ('ChangeAddressTable',
                'Add',
                'AddAll')


class Tab2Handler(unohelper.Base,
                  XContainerWindowEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        try:
            handled = False
            # TODO: During WizardPage initialization the listener must be disabled...
            if method == 'Remove':
                self._manager.removeItem()
                handled = True
            elif method == 'RemoveAll':
                self._manager.removeAllItem()
                handled = True
            return handled
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)

    def getSupportedMethodNames(self):
        return ('Remove',
                'RemoveAll')


class AddressHandler(unohelper.Base,
                     XRowSetListener):
    def __init__(self, manager):
        self._manager = manager

    # XRowSetListener
    def disposing(self, event):
        pass
    def cursorMoved(self, event):
        pass
    def rowChanged(self, event):
        pass
    def rowSetChanged(self, event):
        try:
            self._manager.changeAddressRowSet(event.Source)
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)


class RecipientHandler(unohelper.Base,
                       XRowSetListener):
    def __init__(self, manager):
        self._manager = manager

    # XRowSetListener
    def disposing(self, event):
        pass
    def cursorMoved(self, event):
        pass
    def rowChanged(self, event):
        pass
    def rowSetChanged(self, event):
        try:
            self._manager.changeRecipientRowSet(event.Source)
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)
