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

from com.sun.star.ui.dialogs import XWizardPage

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService

from .mergerview import MergerView
from smtpserver.wizard import getSelectedItems

from smtpserver import logMessage
from smtpserver import getMessage
g_message = 'mergermanager'

import traceback


class MergerManager(unohelper.Base,
                    XWizardPage):
    def __init__(self, ctx, wizard, model, parent, pageid):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        datasources = self._model.getAvailableDataSources()
        self._view = MergerView(ctx, self, parent, datasources)
        self._pageid = pageid
        self._enabled = True

        self._modified = False
        self._index = -1

        datasource = self._model.getDocumentDataSource()
        if datasource in datasources:
            self._view.setPageStep(1)
            # TODO: We must disable the handler "PageHandler" otherwise it activates twice
            #self.disableHandler()
            self._view.selectDataSource(datasource)
            #self.enableHandler()
            #self._model.setDataSource(datasource, self.progress, self.setDataSource)
        else:
            self._view.enableDatasource(True)
        print("PageManager.__init__() 1")

    @property
    def Model(self):
        return self._model

    @property
    def HandlerEnabled(self):
        return self._enabled
    def disableHandler(self):
        self._enabled = False
    def enableHandler(self):
        self._enabled = True

    # XWizardPage
    @property
    def PageId(self):
        return self._pageid
    @property
    def Window(self):
        return self._view.getWindow()

    def activatePage(self):
        try:
            pass
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            logMessage(self._ctx, SEVERE, msg, 'WizardPage', 'activatePage()')
            print("WizardPage.activatePage() %s" % msg)

    def commitPage(self, reason):
        try:
            keys = self._view.getIndexColumns()
            emails = self._view.getEmailColumns()
            recipient = self._view.getSelectedTable()
            #view = self.getView(2)
            # TODO: We can know if "WizardPage2" is already loaded if the view is not None
            #loaded = view is not None
            #address = view.getSelectedAbbressBook() if loaded else None
            #self._model.setRowSet(self.callBack2, loaded, address, recipient, emails, keys)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            logMessage(self._ctx, SEVERE, msg, 'WizardPage1', 'commitPage()')
            print("WizardPage1.commitPage() %s" % msg)

    def canAdvance(self):
        return all((self._view.hasIndex(),
                    self._view.isDataSourceSelected()))

# WizardPage1 methods
    def getPageStep(self, id):
        return self._model.getPageStep(id)

    def changeDataSource(self, datasource):
        self._view.enablePage(False)
        self._view.enableButton(False)
        self._view.setPageStep(1)
        self._model.setDataSource(datasource, self.progress, self.setDataSource)
        #self._wizard.updateTravelUI()

    def progress(self, value):
        self._view.updatePage(value)

    def setDataSource(self, step, queries, tables, email, index, msg):
        if step == 2:
            self._view.setMessageText(msg)
        elif step == 3:
            self._setDataSource(queries, tables, email, index)
            self._view.enablePage(True)
        self._view.setPageStep(step)
        self._wizard.updateTravelUI()

    def changeTables(self, table):
        columns = self._model.getTableColumns(table)
        print("MergerManager.changeTables() 1")
        self._view.initColumns(columns)
        print("MergerManager.changeTables() 2")
        #self._initColumns(columns)
        #self._wizard.updateTravelUI()

    def changeColumns(self, enabled, column):
        print("MergerManager.changeColumns() 1")
        self._view.updateButtonAdd(enabled, column)
        print("MergerManager.changeColumns() 2")
        #self._wizard.updateTravelUI()

    def newDataSource(self):
        frame = self._model.getCurrentDocument().CurrentController.Frame
        service = 'com.sun.star.frame.DispatchHelper'
        dispatcher = createService(self._ctx, service)
        dispatcher.executeDispatch(frame, '.uno:AutoPilotAddressDataSource', '', 0, ())
        #elif tag == 'AddressBook':
        #    dispatcher.executeDispatch(frame, '.uno:AddressBookSource', '', 0, ())
        # TODO: Handler must be disabled because we want just to
        # TODO: update DataSource list and keep the selection...
        datasources = self._model.getAvailableDataSources()
        self.disableHandler()
        self._view.initDataSource(datasources)
        self.enableHandler()

    def changeEmail(self, imax, position):
        self._view.enabledRemoveEmailButton(position != -1)
        self._view.enabledBeforeButton(position > 0)
        self._view.enabledAfterButton(-1 < position < imax)

    def addEmail(self):
        self._view.addEmail()

    def removeEmail(self):
        self._view.removeEmail()

    def moveBefore(self):
        self._modified = True
        self._view.moveBefore()

    def moveAfter(self):
        self._modified = True
        self._view.moveAfter()

    def changeIndex(self, enabled):
        self._view.enabledRemoveIndexButton(enabled)

    def addIndex(self):
        self._view.addIndex()
        self._wizard.updateTravelUI()

    def removeIndex(self):
        self._view.removeIndex()
        self._wizard.updateTravelUI()

    def editQuery(self):
        print("MergerManager.editQuery() ***************************")

    def changeQuery(self):
        print("MergerManager.changeQuery() ***************************")

    def enterQuery(self):
        print("MergerManager.enterQuery() ***************************")

    def addQuery(self):
        print("MergerManager.addQuery() ***************************")

    def removeQuery(self):
        print("MergerManager.removeQuery() ***************************")

# WizardPage1 private methods
    def _setDataSource(self, queries, tables, email, index):
        self._view.initTables(tables)
        # TODO: We must disable the handler "MergerHandler" otherwise it activates twice
        #self.disableHandler()
        #self._view.setTables(table.Name)
        #self.enableHandler()
        self._view.setEmail(email)
        self._view.setIndex(index)
        self._view.initQuery(queries)
        #columns = table.getColumns().getElementNames()
        #self._initColumns(columns)

    def _initColumns(self, columns):
        self._view.initColumns(columns)
        # TODO: We must disable the handler "MergerHandler" otherwise it activates twice
        self.disableHandler()
        self._view.setColumns(0)
        self.enableHandler()
        self._view.updateColumns()









    # TODO: Don't use 'self.View' here because: 'Wizard.getCurrentPage()'
    # TODO: is not yet initialized, and returns the previous page...
    def initPage(self, pageid, window):
        view = MergerView(self._ctx, window)
        self._views[pageid] = view
        if pageid == 1:
            datasources = self.Model.DataSource.getAvailableDataSources()
            view.initPage1(datasources)
            datasource = self.Model.getDocumentDataSource()
            if datasource in datasources:
                view.setPageStep(1)
                # TODO: We must disable the handler "PageHandler" otherwise it activates twice
                self.disableHandler()
                view.selectDataSource(datasource)
                self.enableHandler()
                self.Model.setDataSource(self.progress, datasource, self.callBack1)
            else:
                view.enableDatasource(True)
        elif pageid == 2:
            address = self.Model.DataSource.DataBase.Address
            recipient = self.Model.DataSource.DataBase.Recipient
            view.initGrid(self, address, recipient)
            tables = self.getView(1).getTables()
            columns = self.getView(1).getColumns()
            keys = self.getView(1).getIndexColumns()
            view.initPage2(tables, columns)
            table = self.getView(1).getSelectedTable()
            index = self.Model.DataSource.DataBase.getOrderIndex(columns)
            # TODO: We must disable the handler "PageHandler" otherwise it activates twice
            self.disableHandler()
            view.setPage2(table, index)
            self.enableHandler()
            self.Model.DataSource.initPage2(self.callBack2, keys)




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

    def updateUI1(self, control):
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





