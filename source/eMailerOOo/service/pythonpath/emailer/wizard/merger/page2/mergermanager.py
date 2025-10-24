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

import uno
import unohelper

from com.sun.star.awt import XCallback

from com.sun.star.lang import XComponent

from com.sun.star.ui.dialogs import XWizardPage

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .mergerview import MergerView

from .mergerhandler import Tab1Handler
from .mergerhandler import Tab2Handler

from .mergerhandler import AddressHandler
from .mergerhandler import RecipientHandler

from ....grid import GridSelectionListener

from ....unotool import getArgumentSet
from ....unotool import getStringResource

from ....configuration import g_identifier

import traceback


class MergerManager(unohelper.Base,
                    XWizardPage,
                    XComponent,
                    XCallback):
    def __init__(self, ctx, wizard, model, pageid, parent):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._listeners = []
        self._disabled = False
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'MergerTab')
        tables, title1, title2, message = self._model.getPageInfos(self._resolver)
        self._view = MergerView(ctx, Tab1Handler(self), Tab2Handler(self), parent, tables, title1, title2, message)
        window1, window2 = self._view.getGridWindows()
        self._model.initPage2(window1, window2, self)

    # FIXME: One shot disabler handler
    def isHandlerEnabled(self):
        enabled = True
        if self._disabled:
            self._disabled = enabled = False
        return enabled
    def _disableHandler(self):
        self._disabled = True

# XCallback
    def notify(self, data):
        self._notify(**getArgumentSet(data))

# XComponent
    def dispose(self):
        event = uno.createUnoStruct('com.sun.star.lang.EventObject', self)
        for listener in self._listeners:
            listener.disposing(event)
        self._view.getWindow().dispose()

    def addEventListener(self, listener):
        if listener not in self._listeners:
            self._listeners.append(listener)

    def removeEventListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

# XWizardPage
    @property
    def PageId(self):
        return self._pageid
    @property
    def Window(self):
        return self._view.getWindow()

    def activatePage(self):
        if self._model.hasQueryChanged():
            label = self._model.getQueryLabel(self._resolver)
            self._view.setQueryLabel(label)
        if self._model.hasTablesChanged():
            tables, table = self._model.getQueryTables()
            # FIXME: We must disable the handler "ChangeAddressTable"
            # FIXME: otherwise it activates twice
            self._disableHandler()
            self._view.initTables(tables, table)

    def commitPage(self, reason):
        self._model.resetPendingChanges()
        return True

    def canAdvance(self):
        return self._model.canAdvancePage2()

# MergerManager setter methods
    def setAddressRowSet(self, rowset):
        self._model.setGrid1Data(rowset)

    def resetAddressRowSet(self):
        self._model.resetGrid1Data()

    def setRecipientRowSet(self, rowset):
        self._model.setGrid2Data(rowset)
        count = self._model.getFilterCount()
        self._view.enableAddAll(count == 1)
        self._view.enableRemoveAll(count == 0)
        self._wizard.updateTravelUI()

    def resetRecipientRowSet(self):
        self._model.resetGrid2Data()

    def setAddressTable(self, table):
        self._view.enableAddresTable(False)
        self._model.setAddressTable(table, self.enableAddresTable)

    def enableAddresTable(self, enabled):
        self._view.enableAddresTable(enabled)

    def changeGridSelection(self, index, grid):
        selected = index != -1
        enabled = self._model.hasFilters()
        if grid == 1:
            self._view.enableAdd(selected and enabled)
            if selected:
                self._model.setAddressRecord(index)
        elif grid == 2:
            self._view.enableRemove(selected and enabled)
            if selected:
                self._model.setRecipientRecord(index)

    def addItem(self):
        filters = self._model.getGrid1().getSelectedStructuredFilters()
        self._model.addItem(filters)

    def addAllItem(self):
        table = self._view.getTable()
        self._model.addAllItem(table)

    def removeItem(self):
        filters = self._model.getGrid2().getSelectedStructuredFilters()
        self._model.removeItem(filters)

    def removeAllItem(self):
        self._model.removeAllItem()

    def _notify(self, address, recipient, table):
        self._model.getGrid1().addSelectionListener(GridSelectionListener(self, 1))
        self._model.getGrid2().addSelectionListener(GridSelectionListener(self, 2))
        self._model.addAddressListener(AddressHandler(self))
        self._model.addRecipientListener(RecipientHandler(self))
        # FIXME: We must disable the handler "ChangeAddressBook"
        # FIXME: otherwise it activates twice
        self._disableHandler()
        self._view.setTable(table)


