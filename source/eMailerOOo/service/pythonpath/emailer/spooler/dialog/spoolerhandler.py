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

from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.awt.tab import XTabPageContainerListener

from com.sun.star.frame import XDispatchResultListener

from com.sun.star.util import XModifyListener

from com.sun.star.frame.DispatchResultState import SUCCESS

from com.sun.star.sdbc import XRowSetListener

import traceback


class DialogHandler(unohelper.Base,
                    XDialogEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        try:
            handled = False
            if method == 'ToogleSpooler':
                control = event.Source.Model
                control.Enabled = False
                self._manager.toogleSpooler(control.State)
                handled = True
            elif method == 'ClearLogger':
                self._manager.clearLogger()
                handled = True
            elif method == 'Close':
                self._manager.closeSpooler()
                handled = True
            return handled
        except Exception as e:
            msg = "Error: %s" % traceback.format_exc()
            print(msg)

    def getSupportedMethodNames(self):
        return ('ToogleSpooler',
                'ClearLogger'
                'Close')


class TabHandler(unohelper.Base,
                 XContainerWindowEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, dialog, event, method):
        try:
            handled = False
            if method == 'EmlView':
                self._manager.viewEml()
                handled = True
            elif method == 'ClientView':
                self._manager.viewClient()
                handled = True
            elif method == 'WebView':
                self._manager.viewWeb()
                handled = True
            elif method == 'Add':
                self._manager.addDocument()
                handled = True
            elif method == 'Remove':
                self._manager.removeDocument()
                handled = True
            return handled
        except Exception as e:
            msg = "Error: %s" % traceback.format_exc()
            print(msg)

    def getSupportedMethodNames(self):
        return ('EmlView',
                'ClientView',
                'WebView',
                'Add',
                'Remove')


class TabPageListener(unohelper.Base,
                      XTabPageContainerListener):
    def __init__(self, manager):
        self._manager = manager

    # XTabPageContainerListener
    def tabPageActivated(self, event):
        try:
            self._manager.activateTab(event.TabPageID)
        except Exception as e:
            msg = "Error: %s" % traceback.format_exc()
            print(msg)

    def disposing(self, source):
        pass


class RowSetListener(unohelper.Base,
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
            self._manager.setDataModel(event.Source)
        except Exception as e:
            msg = "Error: %s" % traceback.format_exc()
            print(msg)


class LoggerListener(unohelper.Base,
                     XModifyListener):
    def __init__(self, callback):
        self._callback = callback

    # XModifyListener
    def modified(self, event):
        try:
            self._callback()
        except Exception as e:
            msg = f"Error: {traceback.format_exc()}"
            print(msg)

    def disposing(self, event):
        pass

