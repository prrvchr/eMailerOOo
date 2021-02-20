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


class WindowHandler(unohelper.Base,
                    XContainerWindowEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        handled = False
        # TODO: During WizardPage initialization the listener must be disabled...
        if not self._manager.HandlerEnabled:
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
            self._manager.executeDispatch(event.Source.Model.Tag)
            handled = True
        elif method == 'Move':
            self._manager.moveItem(event.Source)
            handled = True
        elif method == 'Add':
            self._manager.addItem(event.Source.Model.Tag)
            handled = True
        elif method == 'AddAll':
            self._manager.addAllItem()
            handled = True
        elif method == 'Remove':
            self._manager.removeItem(event.Source.Model.Tag)
            handled = True
        elif method == 'RemoveAll':
            self._manager.removeAllItem()
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
        index = -1
        control = event.Source
        tag = control.Model.Tag
        selected = control.hasSelectedRows()
        if tag == 'Recipients' and selected:
            index = control.getSelectedRows()[0]
        self._manager.selectionChanged(tag, selected, index)

    def disposing(self, event):
        pass
