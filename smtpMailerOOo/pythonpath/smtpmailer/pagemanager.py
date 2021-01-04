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

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .pagemodel import PageModel
from .pageview import PageView
from .pagehandler import GridHandler
from .wizardtools import getSelectedItems

from unolib import createService

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
        self._disabled = False
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
    @property
    def Disabled(self):
        return self._disabled
    @Disabled.setter
    def Disabled(self, state):
        self._disabled = state

    # TODO: this method is there for asynchronous calls (ie: callback) because:
    # TODO: 'Wizard.getCurrentPage()' is no longer necessarily valid
    def getView(self, pageid):
        return self._views.get(pageid, None)

    # TODO: Don't use 'self.View' here because: 'Wizard.getCurrentPage()'
    # TODO: is not yet initialized, and returns the previous page...
    def initPage(self, pageid, window):
        view = PageView(self.ctx, window)
        self._views[pageid] = view
        if pageid == 1:
            datasources = self.Model.DataSource.getAvailableDataSources()
            view.initPage1(datasources)
            datasource = self.Model.getDocumentDataSource()
            if datasource in datasources:
                view.setPageStep(1)
                # TODO: We must disable the handler "PageHandler" otherwise it activates twice
                self.Disabled = True
                view.selectDataSource(datasource)
                self.Disabled = False
                self.Model.setDataSource(self.progress, datasource, self.callBack1)
            else:
                view.enableDatasource(True)
        elif pageid == 2:
            handler = GridHandler(self)
            address = self.Model.DataSource.DataBase.Address
            recipient = self.Model.DataSource.DataBase.Recipient
            view.initGrid(address, recipient, handler)
            tables = self.getView(1).getTables()
            columns = self.getView(1).getColumns()
            keys = self.getView(1).getIndexColumns()
            view.initPage2(tables, columns)
            table = self.getView(1).getSelectedTable()
            index = self.Model.DataSource.DataBase.getOrderIndex(columns)
            # TODO: We must disable the handler "PageHandler" otherwise it activates twice
            self.Disabled = True
            view.setPage2(table, index)
            self.Disabled = False
            self.Model.DataSource.initPage2(self.callBack2, keys)

    def progress(self, value):
        if self.Wizard.DialogWindow is not None:
            self.getView(1).updatePage(value)

    def callBack1(self, step, tables, table, emails, keys, msg):
        if self.Wizard.DialogWindow is not None:
            view = self.getView(1)
            if step == 2:
                view.setMessageText(msg)
                view.enableDatasource(True)
            elif step == 3:
                self._initTables(view, tables, table, emails, keys)
                view.enablePage(True)
            view.setPageStep(step)
            self.Wizard.updateTravelUI()

    def callBack2(self):
        if self.Wizard.DialogWindow is not None:
            view = self.getView(2)
            view.updateControlByTag('Addresses')
            view.updateControlByTag('Recipients')

    def selectionChanged(self, tag, selected, index):
        if tag == 'Addresses':
            self.View.updateAddRecipient(selected)
        elif tag == 'Recipients':
            self.View.updateRemoveRecipient(selected)
            if selected and index != self._index:
                self.Model.setDocumentRecord(index)
                self._index = index

    def activatePage(self, pageid):
        pass

    def canAdvancePage(self, pageid):
        advance = False
        if pageid == 1:
            advance = self.View.canAdvancePage1()
        elif pageid == 2:
            advance = self.Model.DataSource.DataBase._recipient.RowCount != 0
        elif pageid == 3:
            pass
        return advance

    def commitPage(self, pageid, reason):
        if pageid == 1:
            keys = self.View.getIndexColumns()
            emails = self.View.getEmailColumns()
            recipient = self.View.getSelectedTable()
            view = self.getView(2)
            # TODO: We can know if "WizardPage2" is already loaded if the view is not None
            loaded = view is not None
            address = view.getSelectedAbbressBook() if loaded else None
            self.Model.setRowSet(self.callBack2, loaded, address, recipient, emails, keys)
        elif pageid == 2:
            pass
        elif pageid == 3:
            pass
        return True

    def updateUI(self, control):
        try:
            tag = control.Model.Tag
            if tag == 'DataSource':
                datasource = control.getSelectedItem()
                view = self.getView(1)
                view.enablePage(False)
                view.enableButton(False)
                view.setPageStep(1)
                self.Model.setDataSource(self.progress, datasource, self.callBack1)
            elif tag == 'AddressBook':
                table = control.getSelectedItem()
                view = self.getView(1)
                emails = view.getEmailColumns()
                keys = view.getIndexColumns()
                self.Model.executeAddress(self.callBack2, emails, table, keys)
                #columns = self.Model.DataSource.DataBase.getTableColumns(table)
                #if self.View.getOrderColumns() != columns:
                #    print("PageManager.updateUI() ERROR ***********")
                #    self.View.setOrder(self.Model.DataSource.DataBase.getOrderIndex())
                #self.View.refreshGridButton()
            elif tag == 'Columns':
                self.View.updateControl(control)
            elif tag == 'OrderColumns':
                view = self.getView(1)
                keys = view.getIndexColumns()
                recipient = view.getSelectedTable()
                address = self.View.getSelectedAbbressBook()
                order = getSelectedItems(control)
                self.Model.executeRowSet(self.callBack2, address, recipient, order, keys)
                #self.View.refreshGridButton()
            elif tag == 'EmailAddress':
                self.View.updateControl(control)
            elif tag == 'PrimaryKey':
                self.View.updateControl(control)
            elif tag == 'Tables':
                table = control.getSelectedItem()
                columns = self.Model.getTableColumns(table)
                print("PageManager.updateUI() %s ***************************" % (columns, ))
                self._initColumns(self.View, columns)
        except Exception as e:
            print("PageManager.updateUI() ERROR: %s - %s" % (e, traceback.print_exc()))

    def updateTravelUI(self):
        self.Wizard.updateTravelUI()

    def setPageTitle(self, pageid):
        title = self.getView(pageid).getPageTitle(self.Model, pageid)
        self.Wizard.setTitle(title)

    def executeDispatch(self, tag):
        frame = self.Model.getCurrentDocument().CurrentController.Frame
        dispatcher = createService(self.ctx, 'com.sun.star.frame.DispatchHelper')
        if tag == 'DataSource':
            dispatcher.executeDispatch(frame, '.uno:AutoPilotAddressDataSource', '', 0, ())
        elif tag == 'AddressBook':
            dispatcher.executeDispatch(frame, '.uno:AddressBookSource', '', 0, ())
        # TODO: Handler must be disabled because we want just to
        # TODO: update DataSource list and keep the selection...
        self.Disabled = True
        self.View.initDataSource(self.Model.DataSource.getAvailableDataSources())
        self.Disabled = False

    def addItem(self, tag):
        try:
            self._modified = True
            if tag == 'EmailAddress':
                self.View.addEmailAdress()
            elif tag == 'PrimaryKey':
                self.View.addPrimaryKey()
                self.Wizard.updateTravelUI()
            elif tag == 'Recipient':
                indexes = self.getView(1).getIndexColumns()
                recipients = self.Model.DataSource.DataBase.getRecipientFilters(indexes)
                rows = self.View.getSelectedAddress()
                filters = self.Model.DataSource.DataBase.getAddressFilters(indexes, rows, recipients)
                self.Model.executeRecipient(self.callBack2, indexes, recipients + filters)
        except Exception as e:
            print("PageManager.addItem() ERROR: %s - %s" % (e, traceback.print_exc()))

    def addAllItem(self):
        try:
            self._modified = True
            indexes = self.getView(1).getIndexColumns()
            recipients = self.Model.DataSource.DataBase.getRecipientFilters(indexes)
            rows = self.Model.DataSource.DataBase.getAddressRows()
            filters = self.Model.DataSource.DataBase.getAddressFilters(indexes, rows, recipients)
            self.Model.executeRecipient(self.callBack2, indexes, recipients + filters)
        except Exception as e:
            print("PageManager.addAllItem() ERROR: %s - %s" % (e, traceback.print_exc()))

    def removeItem(self, tag):
        try:
            self._modified = True
            if tag == 'EmailAddress':
                self.View.removeEmailAdress()
            elif tag == 'PrimaryKey':
                self.View.removePrimaryKey()
                self.Wizard.updateTravelUI()
            elif tag == 'Recipient':
                rows = self.View.getSelectedRecipients()
                indexes = self.getView(1).getIndexColumns()
                filters = self.Model.DataSource.DataBase.getRecipientFilters(indexes, rows)
                self.Model.executeRecipient(self.callBack2, indexes, filters)
        except Exception as e:
            print("PageManager.removeItem() ERROR: %s - %s" % (e, traceback.print_exc()))

    def removeAllItem(self):
        try:
            self._modified = True
            self.View.deselectAllRecipients()
            indexes = self.getView(1).getIndexColumns()
            self.Model.executeRecipient(self.callBack2, indexes)
        except Exception as e:
            print("PageManager.removeAllItem() ERROR: %s - %s" % (e, traceback.print_exc()))

    def moveItem(self, control):
        self._modified = True
        self.View.moveEmailAdress(control.Model.Tag)

    def _initTables(self, view, tables, table, emails, keys):
        view.initTables(tables)
        # TODO: We must disable the handler "PageHandler" otherwise it activates twice
        self.Disabled = True
        view.setTables(table.Name)
        self.Disabled = False
        view.setEmailAddress(emails)
        view.setPrimaryKey(keys)
        columns = table.getColumns().getElementNames()
        self._initColumns(view, columns)

    def _initColumns(self, view, columns):
        view.initColumns(columns)
        # TODO: We must disable the handler "PageHandler" otherwise it activates twice
        self.Disabled = True
        view.setColumns(0)
        self.Disabled = False
        view.updateControlByTag('Columns')

