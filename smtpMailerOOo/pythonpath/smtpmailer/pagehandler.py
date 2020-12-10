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
from com.sun.star.awt.grid import XGridSelectionListener

import traceback


class PageHandler(unohelper.Base,
                  XContainerWindowEventHandler):
    def __init__(self, manager):
        self._manager = manager
        self._disabled = False

    @property
    def Manager(self):
        return self._manager

    def disable(self):
        self._disabled = True

    def enable(self):
        self._disabled = False

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        handled = False
        # TODO: During WizardPage initialization the listener must be disabled...
        if self._disabled:
            handled = True
        elif method == 'StateChange':
            self._manager.updateUI(event.Source)
            handled = True
        elif method == 'SettingChanged':
            self._manager.updateUI(event.Source)
            handled = True
        elif method == 'ColumnChanged':
            print("PageHandler.callHandlerMethod() ColumnChanged ***************")
            self._manager.updateUI(event.Source)
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
            self._manager.moveItem(event.Source)
            handled = True
        elif method == 'Add':
            self._manager.addItem(event.Source)
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
            self._manager.removeItem(event.Source)
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


class GridHandler(unohelper.Base,
                  XGridSelectionListener):
    def __init__(self, manager):
        self._manager = manager

    # XGridSelectionListener
    def selectionChanged(self, event):
        self._manager.selectionChanged(event.Source)

    def disposing(self, event):
        pass
