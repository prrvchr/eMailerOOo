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
from com.sun.star.awt.grid import XGridSelectionListener
from com.sun.star.frame import XDispatchResultListener

from com.sun.star.frame.DispatchResultState import SUCCESS

import traceback


class DialogHandler(unohelper.Base,
                    XDialogEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'ToogleSpooler':
            self._manager.toogleSpooler()
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('ToogleSpooler',)


class Page1Handler(unohelper.Base,
                   XContainerWindowEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'ChangeColumn':
            columns = event.Source.getItems()
            self._manager.changeColumn(columns)
            handled = True
        elif method == 'ChangeOrder':
            orders = event.Source.getItems()
            self._manager.changeOrder(orders)
            handled = True
        elif method == 'Add':
            self._manager.addDocument()
            handled = True
        elif method == 'Remove':
            self._manager.removeDocument()
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('ChangeColumn',
                'ChangeOrder',
                'Add',
                'Remove')


class Page2Handler(unohelper.Base,
                   XContainerWindowEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'AddAttachment':
            self._manager.addAttachments()
            handled = True
        elif method == 'RemoveAttachment':
            self._manager.removeAttachments()
            handled = True
        elif method == 'ChangeAttachments':
            enabled = event.Source.getSelectedItemPos() != -1
            self._manager.enableRemoveAttachments(enabled)
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('AddAttachment',
                'RemoveAttachment',
                'ChangeAttachments')


class GridHandler(unohelper.Base,
                  XGridSelectionListener):
    def __init__(self, manager):
        self._manager = manager

    # XGridSelectionListener
    def selectionChanged(self, event):
        enabled = event.Source.hasSelectedRows()
        self._manager.toogleRemove(enabled)

    def disposing(self, source):
        pass


class DispatchListener(unohelper.Base,
                       XDispatchResultListener):
    def __init__(self, manager):
        self._manager = manager

    # XDispatchResultListener
    def dispatchFinished(self, notification):
        if notification.State == SUCCESS:
            self._manager.addJob(notification.Result)

    def disposing(self, source):
        pass
