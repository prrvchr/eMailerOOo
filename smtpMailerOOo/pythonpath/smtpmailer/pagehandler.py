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

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        handled = False
        # TODO: During ListBox initializing the listener must be disabled...
        if not self._enabled:
            handled = True
        elif method == 'StateChange':
            #handled = self._updateUI(window, event)
            handled = True
        elif method == 'SettingChanged':
            #control = event.Source
            #handled = self._changeSetting(window, control)
            handled = True
        elif method == 'ColumnChanged':
            #control = event.Source
            #handled = self._changeColumn(window, control)
            handled = True
        elif method == 'OutputChanged':
            #control = event.Source
            #handled = self._outputChanged(window, control)
            handled = True
        elif method == 'Dispatch':
            #control = event.Source
            #handled = self._executeDispatch(control)
            handled = True
            #self._updateControl(window, control)
        elif method == 'Move':
            #control = event.Source
            #handled = self._moveItem(window, control)
            handled = True
        elif method == 'Add':
            #control = event.Source
            #handled = self._addItem(window, control)
            handled = True
        elif method == 'AddAll':
            #self._modified = True
            #grid = window.getControl('GridControl2')
            #recipients = self._getRecipientFilters()
            #rows = range(self._address.RowCount)
            #filters = self._getAddressFilters(rows, recipients)
            #handled = self._rowRecipientExecute(recipients + filters)
            #self._updateControl(window, grid)
            handled = True
        elif method == 'Remove':
            #control = event.Source
            #handled = self._removeItem(window, control)
            handled = True
        elif method == 'RemoveAll':
            #self._modified = True
            #grid = window.getControl('GridControl2')
            #grid.deselectAllRows()
            #handled = self._rowRecipientExecute()
            #self._updateControl(window, grid)
            handled = True
        self._manager.updateTravelUI()
        return handled

    def getSupportedMethodNames(self):
        return ('StateChange',
                'Dispatch',
                'Add',
                'AddAll',
                'Remove',
                'RemoveAll',
                'OutputChanged',
                'SettingChanged',
                'ColumnChanged',
                'Move')
