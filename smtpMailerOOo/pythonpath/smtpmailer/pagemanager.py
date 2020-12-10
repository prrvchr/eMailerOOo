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

from com.sun.star.uno import Exception as UnoException
from com.sun.star.beans import PropertyValue

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .pagemodel import PageModel
from .pageview import PageView
from .pagehandler import GridHandler
from .wizardtools import getSelectedItems

from unolib import createService
from unolib import getUrl

from .logger import logMessage
from .logger import getMessage
g_message = 'pagemanager'


import traceback


class PageManager(unohelper.Base):
    def __init__(self, ctx, wizard, model=None):
        self.ctx = ctx
        self._wizard = wizard
        self._model = PageModel(self.ctx) if model is None else model
        self._views = {}
        self._modified = False
        self._index = -1
        print("PageManager.__init__() 1")

    @property
    def Model(self):
        return self._model
    @property
    def Wizard(self):
        return self._wizard
    @property
    def View(self):
        pageid = self.Wizard.getCurrentPage().PageId
        return self.getView(pageid)

    # TODO: this method is there for asynchronous calls (ie: callback) because:
    # TODO: 'Wizard.getCurrentPage()' is no longer necessarily valid
    def getView(self, pageid):
        if pageid in self._views:
            return self._views[pageid]
        print("PageManager.getView ERROR **************************")
        return None

    # TODO: Don't use 'self.View' here because: 'Wizard.getCurrentPage()'
    # TODO: is not yet initialized, and returns the previous page...
    def initPage(self, pageid, window):
        print("PageManager.initPage() 1")
        view = PageView(self.ctx, window)
        self._views[pageid] = view
        if pageid == 1:
            print("PageManager.initPage() 2")
            if view.initPage1(self.Model):
                print("PageManager.initPage() 3")
                self.Wizard.updateTravelUI()
        elif pageid == 2:
            print("PageManager.initPage() 4")
            handler = GridHandler(self)
            view.initPage2(self.Model, handler)
            print("PageManager.initPage() 5")
            #self._handler.addRefreshListener(self)
            self.Model._recipient.execute()
            print("PageManager.initPage() 6")
            table = view.initAddressBook(self.Model)
            view.initOrderColumn(self.Model)
            print("PageManager.initPage() 7")
            self.Model.setAddressBook(view, table)
            view.refreshGridButton()
        print("PageManager.initPage() 8")

    def getWizard(self):
        print("PageManager.getWizard() ERROR ***************************************")
        return self.Wizard

    def selectionChanged(self, control):
        tag = control.Model.Tag
        enabled = control.hasSelectedRows()
        if tag == 'Addresses':
            self.View.updateAddRecipient(enabled)
        elif tag == 'Recipients':
            self.View.updateRemoveRecipient(enabled)
            if enabled:
                index = control.getSelectedRows()[0]
                if index != self._index:
                    self._setDocumentRecord(index)

    def activatePage(self, pageid):
        if pageid == 1:
            pass

    def canAdvancePage(self, pageid):
        advance = False
        if pageid == 1:
            advance = all((self.Model.Connection is not None,
                           self.View.isPrimaryKeySet()))
        elif pageid == 2:
            advance = self.Model._recipient.RowCount != 0
        elif pageid == 3:
            pass
        return advance

    def commitPage(self, pageid, reason):
        if pageid == 1:
            pass
        elif pageid == 2:
            pass
        elif pageid == 3:
            pass
        return True

    def updateUI(self, control):
        try:
            tag = control.Model.Tag
            if tag == 'DataSource':
                self.Model.setDataSource(self.View, control.getSelectedItem())
                self.Wizard.updateTravelUI()
            elif tag == 'AddressBook':
                if self.Model.setAddressBook(self.View, control.getSelectedItem()):
                    self.View.initOrderColumn(self.Model)
                self.View.refreshGridButton()
            elif tag == 'Columns':
                self.View.updateControl(control)
            elif tag == 'OrderColumns':
                self.Model.setOrderColumn(self.View, getSelectedItems(control))
            elif tag == 'EmailAddress':
                self.View.updateControl(control)
            elif tag == 'PrimaryKey':
                self.View.updateControl(control)
            elif tag == 'Columns1':
                pass
                #self._changeColumns(window, getSelectedItems(control))
            elif tag == 'Custom':
                self.View.updateTables(control.Model.State)
            elif tag == 'Tables':
                self.Model.initColumns(self.View, control.getSelectedItem())
        except Exception as e:
            print("PageManager.updateUI() ERROR: %s - %s" % (e, traceback.print_exc()))

    def updateTravelUI(self):
        self.Wizard.updateTravelUI()

    def setPageTitle(self, pageid):
        title = self.getView(pageid).getPageTitle(self.Model, pageid)
        self.Wizard.setTitle(title)

    def addItem(self, control):
        self._modified = True
        tag = control.Model.Tag
        if tag == 'EmailAddress':
            self.View.addEmailAdress()
        elif tag == 'PrimaryKey':
            self.View.addPrimaryKey()
            self.Wizard.updateTravelUI()
        elif tag == 'Recipient':
            grid = window.getControl('GridControl2')
            recipients = self._getRecipientFilters()
            rows = window.getControl('GridControl1').getSelectedRows()
            filters = self._getAddressFilters(rows, recipients)
            handled = self._rowRecipientExecute(recipients + filters)
            self._updateControl(window, grid)

    def removeItem(self, control):
        self._modified = True
        tag = control.Model.Tag
        if tag == 'EmailAddress':
            self.View.removeEmailAdress()
        elif tag == 'PrimaryKey':
            self.View.removePrimaryKey()
            self.Wizard.updateTravelUI()
        elif tag == 'Recipient':
            grid = window.getControl('GridControl2')
            filters = self._getRecipientFilters(grid.getSelectedRows())
            grid.deselectAllRows()
            handled = self._rowRecipientExecute(filters)
            self._updateControl(window, grid)

    def moveItem(self, control):
        self._modified = True
        self.View.moveEmailAdress(control.Model.Tag)

    def _setDocumentRecord(self, index):
        try:
            dispatch = None
            document = createService(self.ctx, 'com.sun.star.frame.Desktop').CurrentComponent
            frame = document.getCurrentController().Frame
            flag = uno.getConstantByName('com.sun.star.frame.FrameSearchFlag.SELF')
            if document.supportsService('com.sun.star.text.TextDocument'):
                url = getUrl('.uno:DataSourceBrowser/InsertContent')
                dispatch = frame.queryDispatch(url, '_self', flag)
            elif document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
                url = getUrl('.uno:DataSourceBrowser/InsertColumns')
                dispatch = frame.queryDispatch(url, '_self', flag)
            if dispatch is not None:
                dispatch.dispatch(url, self._getDataDescriptor(index + 1))
                self._index = index
        except Exception as e:
            print("PageManager._setDocumentRecord() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _getDataDescriptor(self, row):
        descriptor = []
        direct = uno.Enum('com.sun.star.beans.PropertyState', 'DIRECT_VALUE')
        recipient = self.Model._recipient
        connection = recipient.ActiveConnection
        descriptor.append(PropertyValue('DataSourceName', -1, recipient.DataSourceName, direct))
        descriptor.append(PropertyValue('ActiveConnection', -1, connection, direct))
        descriptor.append(PropertyValue('Command', -1, recipient.Command, direct))
        descriptor.append(PropertyValue('CommandType', -1, recipient.CommandType, direct))
        descriptor.append(PropertyValue('Cursor', -1, recipient, direct))
        descriptor.append(PropertyValue('Selection', -1, (row, ), direct))
        descriptor.append(PropertyValue('BookmarkSelection', -1, False, direct))
        return tuple(descriptor)
