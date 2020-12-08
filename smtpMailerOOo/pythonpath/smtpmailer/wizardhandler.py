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

from com.sun.star.util import XRefreshable
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.beans import PropertyValue
from com.sun.star.util import URL
from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE
from com.sun.star.lang import WrappedTargetException
from com.sun.star.ucb import XCommandEnvironment

from unolib import PropertySet
from unolib import createService
from unolib import getProperty
from unolib import getStringResource
from unolib import getPropertyValueSet

from .griddatamodel import GridDataModel
from .dbtools import getValueFromResult
from .dbqueries import getSqlQuery
from .wizardtools import getRowSetOrders
from .wizardtools import getSelectedItems
from .wizardtools import getStringItemList

from .configuration import g_identifier
from .configuration import g_extension
from .configuration import g_column_index
from .configuration import g_column_filters
from .configuration import g_table_index
from .configuration import g_fetchsize

from .logger import logMessage

import traceback


class WizardHandler(unohelper.Base,
                    PropertySet,
                    XContainerWindowEventHandler,
                    XRefreshable):
    def __init__(self, ctx, wizard):
        self.ctx = ctx
        self._wizard = wizard
        #self._dbcontext = createService(self.ctx, 'com.sun.star.sdb.DatabaseContext')
        self._listeners = []
        self._disabled = False
        self._modified = False
        self._address = createService(self.ctx, 'com.sun.star.sdb.RowSet')
        self._address.CommandType = TABLE
        self._address.FetchSize = g_fetchsize
        self._recipient = createService(self.ctx, 'com.sun.star.sdb.RowSet')
        self._recipient.CommandType = TABLE
        self._recipient.FetchSize = g_fetchsize
        self._statement = None
        self._table = None
        self._database = None
        self._quey = None
        self._emailcolumns = self._indexcolumns = ()
        self.index = -1

    @property
    def DataSources(self):
        dbcontext = createService(self.ctx, 'com.sun.star.sdb.DatabaseContext')
        return dbcontext.getElementNames()
    @property
    def Connection(self):
        if self._statement is not None:
            return self._statement.getConnection()
        return None
    @property
    def TableNames(self):
        if self.Connection is not None:
            return self.Connection.getTables().getElementNames()
        return ()
    @property
    def ColumnNames(self):
        if self._table is not None:
            return self._table.getColumns().getElementNames()
        return ()

    # XRefreshable
    def refresh(self):
        pass
    def addRefreshListener(self, listener):
        if listener not in self._listeners:
            self._listeners.append(listener)
    def removeRefreshListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)
    def _refresh(self, control):
        event = uno.createUnoStruct('com.sun.star.lang.EventObject', control)
        for listener in self._listeners:
            listener.refreshed(event)

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        try:
            # TODO: During ListBox initializing the listener must be disabled...
            if self._disabled:
                return True
            handled = False
            if method == 'StateChange':
                handled = self._updateUI(window, event)
            elif method == 'SettingChanged':
                control = event.Source
                handled = self._changeSetting(window, control)
            elif method == 'ColumnChanged':
                control = event.Source
                handled = self._changeColumn(window, control)
            elif method == 'OutputChanged':
                control = event.Source
                handled = self._outputChanged(window, control)
            elif method == 'Dispatch':
                control = event.Source
                handled = self._executeDispatch(control)
                #self._updateControl(window, control)
            elif method == 'Move':
                control = event.Source
                handled = self._moveItem(window, control)
            elif method == 'Add':
                control = event.Source
                handled = self._addItem(window, control)
            elif method == 'AddAll':
                self._modified = True
                grid = window.getControl('GridControl2')
                recipients = self._getRecipientFilters()
                rows = range(self._address.RowCount)
                filters = self._getAddressFilters(rows, recipients)
                handled = self._rowRecipientExecute(recipients + filters)
                self._updateControl(window, grid)
            elif method == 'Remove':
                control = event.Source
                handled = self._removeItem(window, control)
            elif method == 'RemoveAll':
                self._modified = True
                grid = window.getControl('GridControl2')
                grid.deselectAllRows()
                handled = self._rowRecipientExecute()
                self._updateControl(window, grid)
            self._wizard.updateTravelUI()
            return handled
        except Exception as e:
            print("WizardHandler.callHandlerMethod() ERROR: %s - %s" % (e, traceback.print_exc()))
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

    def _changeColumn(self, window, control):
        self._updateControl(window, control)
        return True

    def _moveItem(self, window, control):
        self._modified = True
        tag = control.Model.Tag
        listbox = window.getControl('ListBox4')
        index = listbox.getSelectedItemPos()
        column = listbox.getSelectedItem()
        columns = self._removeColumn(listbox, column, True)
        if tag == 'Before':
            index -= 1
        elif tag == 'After':
            index += 1
        listbox.Model.StringItemList = self._addColumn(columns, column, index)
        listbox.selectItemPos(index, True)
        return True

    def _addItem(self, window, control):
        self._modified = True
        tag = control.Model.Tag
        if tag == 'EmailAddress':
            window.getControl('CommandButton2').Model.Enabled = False
            window.getControl('CommandButton3').Model.Enabled = False
            window.getControl('CommandButton4').Model.Enabled = False
            window.getControl('CommandButton5').Model.Enabled = False
            listbox = window.getControl('ListBox4')
            handled = self._addItemColumn(window, listbox)
        elif tag == 'PrimaryKey':
            window.getControl('CommandButton6').Model.Enabled = False
            window.getControl('CommandButton7').Model.Enabled = False
            listbox = window.getControl('ListBox5')
            handled = self._addItemColumn(window, listbox)
        elif tag == 'Recipient':
            grid = window.getControl('GridControl2')
            recipients = self._getRecipientFilters()
            rows = window.getControl('GridControl1').getSelectedRows()
            filters = self._getAddressFilters(rows, recipients)
            handled = self._rowRecipientExecute(recipients + filters)
            self._updateControl(window, grid)
        return handled

    def _removeItem(self, window, control):
        self._modified = True
        tag = control.Model.Tag
        if tag == 'EmailAddress':
            window.getControl('CommandButton3').Model.Enabled = False
            window.getControl('CommandButton4').Model.Enabled = False
            window.getControl('CommandButton5').Model.Enabled = False
            listbox = window.getControl('ListBox4')
            listbox.Model.StringItemList = self._removeColumn(listbox, listbox.getSelectedItem())
            button = window.getControl('CommandButton2')
            button.Model.Enabled = self._canAdd(window.getControl('ListBox2'), listbox)
            handled = True
        elif tag == 'PrimaryKey':
            window.getControl('CommandButton7').Model.Enabled = False
            listbox = window.getControl('ListBox5')
            listbox.Model.StringItemList = self._removeColumn(listbox, listbox.getSelectedItem())
            button = window.getControl('CommandButton6')
            button.Model.Enabled = self._canAdd(window.getControl('ListBox2'), listbox)
            handled = True
        elif tag == 'Recipient':
            grid = window.getControl('GridControl2')
            filters = self._getRecipientFilters(grid.getSelectedRows())
            grid.deselectAllRows()
            handled = self._rowRecipientExecute(filters)
            self._updateControl(window, grid)
        return handled

    def _addItemColumn(self, window, control):
        column = window.getControl('ListBox2').getSelectedItem()
        columns = list(getStringItemList(control))
        control.Model.StringItemList = self._addColumn(columns, column)
        return True

    def _addColumn(self, columns, column, index=-1):
        if index != -1:
            columns.insert(index, column)
        else:
            columns.append(column)
        return tuple(columns)

    def _removeColumn(self, control, column, mutable=False):
        columns = list(getStringItemList(control))
        columns.remove(column)
        return columns if mutable else tuple(columns)

    def getDocumentDataSource(self):
        setting = 'com.sun.star.document.Settings'
        document = createService(self.ctx, 'com.sun.star.frame.Desktop').CurrentComponent
        if document.supportsService('com.sun.star.text.TextDocument'):
            return document.createInstance(setting).CurrentDatabaseDataSource
        return ''

    def setDocumentRecord(self, index):
        try:
            dispatch = None
            document = createService(self.ctx, 'com.sun.star.frame.Desktop').CurrentComponent
            frame = document.getCurrentController().Frame
            flag = uno.getConstantByName('com.sun.star.frame.FrameSearchFlag.SELF')
            if document.supportsService('com.sun.star.text.TextDocument'):
                url = self._getUrl('.uno:DataSourceBrowser/InsertContent')
                dispatch = frame.queryDispatch(url, '_self', flag)
            elif document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
                url = self._getUrl('.uno:DataSourceBrowser/InsertColumns')
                dispatch = frame.queryDispatch(url, '_self', flag)
            if dispatch is not None:
                dispatch.dispatch(url, self._getDataDescriptor(index + 1))
            self.index = index
        except Exception as e:
            print("WizardHandler.setDocumentRecord() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _changeSetting(self, window, control):
        tag = control.Model.Tag
        if tag == 'Tables':
            self._initColumnsSetting(window, control.getSelectedItem())
        elif tag == 'Columns':
            self._updateControl(window, control)
        elif tag == 'Custom':
            window.getControl('ListBox3').Model.Enabled = bool(control.Model.State)
        return True

    def _initColumnsSetting(self, window, name=None):
        if name is None:
            table = self.Connection.getTables().getByIndex(0)
        else:
            table = self.Connection.getTables().getByName(name)
        self._recipient.Command = table.Name
        columns = table.getColumns().getElementNames()
        listbox = window.getControl('ListBox2')
        listbox.Model.StringItemList = columns
        listbox.selectItemPos(0, True)

    def _getUrl(self, complete):
        url = URL()
        url.Complete = complete
        dummy, url = createService(self.ctx, 'com.sun.star.util.URLTransformer').parseStrict(url)
        return url

    def _getDataDescriptor(self, row):
        descriptor = []
        direct = uno.Enum('com.sun.star.beans.PropertyState', 'DIRECT_VALUE')
        connection = self._recipient.ActiveConnection
        descriptor.append(PropertyValue('DataSourceName', -1, self._recipient.DataSourceName, direct))
        descriptor.append(PropertyValue('ActiveConnection', -1, connection, direct))
        descriptor.append(PropertyValue('Command', -1, self._recipient.Command, direct))
        descriptor.append(PropertyValue('CommandType', -1, self._recipient.CommandType, direct))
        descriptor.append(PropertyValue('Cursor', -1, self._recipient, direct))
        descriptor.append(PropertyValue('Selection', -1, (row, ), direct))
        descriptor.append(PropertyValue('BookmarkSelection', -1, False, direct))
        return tuple(descriptor)

    def _outputChanged(self, window, control):
        tag = control.Model.Tag
        print("WizardHandler._outputChanged() %s" % tag)
        return True

    def _updateUI(self, window, event):
        try:
            control = event.Source
            tag = control.Model.Tag
            handled = False
            if tag == 'DataSource':
                handled = self._changeDataSource(window, control.getSelectedItem())
            elif tag == 'AddressBook':
                handled = self._changeAddressBook(window, control.getSelectedItem())
            elif tag == 'Columns':
                handled = self._changeColumns(window, getSelectedItems(control))
            return handled
        except Exception as e:
            print("WizardHandler._updateUI() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _changeDataSource(self, window, datasource):
        print("WizardHandler._updateUI() DataSource 1")
        database = self._getDatabase(datasource)
        if database is not None and not database.IsPasswordRequired:
            connection = database.getConnection('', '')
            self._database = database
            self._statement = connection.createStatement()
            self._query = self._getQueryComposer()
            self._setRowSet(datasource)
            print("WizardHandler._updateUI() DataSource 2")
            self._initSetting(window)
            self._updateControl(window, window.getControl('ListBox2'))
            self._updateControl(window, window.getControl('ListBox4'))
            self._updateControl(window, window.getControl('ListBox5'))
            print("WizardHandler._updateUI() DataSource 3")
            self._refresh(window.getControl('ListBox1'))
            print("WizardHandler._updateUI() DataSource 4")
            self._wizard.updateTravelUI()
            print("WizardHandler._updateUI() DataSource 5")
        else:
            self._table = None
            self._database = None
            self._statement = None
        return True

    def _getQueryComposer(self):
        composer = self.Connection.createInstance('com.sun.star.sdb.SingleSelectQueryComposer')
        query = self._getQuery(False)
        composer.setQuery(query.Command)
        self._address.Command = query.UpdateTableName
        print("WizardHandler._getQueryComposer() %s - %s" % (query.UpdateTableName, query.Command))
        #mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
        #mri.inspect(composer)
        return composer

    def _setRowSet(self, datasource):
        self._address.DataSourceName = self._recipient.DataSourceName = datasource
        self._address.Order = self._recipient.Order = self._query.getOrder()
        self._recipient.Command = self._query.getTables().getByIndex(0).Name
        self._recipient.Filter = self._query.getFilter()
        print("WizardHandler._setRowSet() %s - %s" % (self._query.ElementaryQuery, self._query.getFilter()))
        self._recipient.ApplyFilter = True
        self._recipient.execute()

    def _changeAddressBook(self, window, table):
        print("WizardHandler._updateUI() AddressBook 1")
        self._table = self.Connection.getTables().getByName(table)
        #self._address.UpdateTableName = table
        self._address.Command = table
        self._address.Filter = self._getFilter(True)
        self._address.ApplyFilter = True
        self._address.execute()
        self._refreshColumns(window.getControl('ListBox2'))
        self._refreshButton(window)
        print("WizardHandler._updateUI() AddressBook 2")
        return True

    def _changeColumns(self, window, columns):
        print("WizardHandler._updateUI() Columns 1")
        # TODO: During ListBox initializing the listener must be disabled...
        if self._disabled:
            return True
        self._modified = True
        # TODO: XRowset.Order should be treated as a column stack
        # TODO: where adding is done at the end and removing will keep order
        orders = getRowSetOrders(self._recipient)
        print("WizardHandler._updateUI() Columns 2: %s - %s" % (orders, columns))
        for order in reversed(orders):
            if order not in columns:
                orders.remove(order)
        for column in columns:
            if column not in orders:
                orders.append(column)
        order = '"%s"' % '", "'.join(orders) if len(orders) else ''
        print("WizardHandler._updateUI() Columns 3: %s" % order)
        self._query.setOrder(order)
        self._setRowSetOrder()
        self._refreshButton(window)
        print("WizardHandler._updateUI() Columns 4")
        return True

    def _refreshButton(self, window):
        self._updateControl(window, window.getControl('GridControl1'))
        self._updateControl(window, window.getControl('GridControl2'))

    def _setRowSetOrder(self):
        print("WizardHandler._setRowSetOrder() 1")
        self._recipient.Order = self._address.Order = self._query.getOrder()
        print("WizardHandler._setRowSetOrder() 2")
        self._address.execute()
        print("WizardHandler._setRowSetOrder() 3")
        self._recipient.execute()

    def _initSetting(self, window):
        doc, form = self._getForm(False)
        self._initTableSetting(doc, window.getControl('ListBox3'), 'PrimaryTable')
        control = window.getControl('ListBox4')
        self._emailcolumns = self._initColumnSetting(doc, control, 'EmailColumns')
        control = window.getControl('ListBox5')
        self._indexcolumns = self._initColumnSetting(doc, control, 'IndexColumns')
        if form is not None:
            form.close()

    def _initTableSetting(self, document, control, property):
        control.Model.StringItemList = self.TableNames
        table = None if document is None else self._getDocumentUserProperty(document, property)
        if table is None:
            control.selectItemPos(0, True)
        else:
            control.selectItem(table, True)

    def _initColumnSetting(self, document, control, property):
        value = None if document is None else self._getDocumentUserProperty(document, property)
        items = () if value is None else tuple(value.split(','))
        control.Model.StringItemList = items
        return items

    def refreshTables(self, control):
        tables = self.TableNames
        control.Model.StringItemList = ()
        control.Model.StringItemList = tables
        table = self._address.Command
        table = tables[0] if len(tables) != 0 and table == '' else table
        control.selectItem(table, True)

    def _refreshColumns(self, control):
        self._disabled = True
        control.Model.StringItemList = ()
        control.Model.StringItemList = self.ColumnNames
        columns = self._getOrderIndex()
        control.selectItemsPos(columns, True)
        self._disabled = False

    def _executeDispatch(self, control):
        tag = control.Model.Tag
        document = createService(self.ctx, 'com.sun.star.frame.Desktop').CurrentComponent
        frame = document.CurrentController.Frame
        dispatcher = createService(self.ctx, 'com.sun.star.frame.DispatchHelper')
        if tag == 'DataSource':
            dispatcher.executeDispatch(frame, '.uno:AutoPilotAddressDataSource', '', 0, ())
        elif tag == 'AddressBook':
            dispatcher.executeDispatch(frame, '.uno:AddressBookSource', '', 0, ())
        return True

    def _rowRecipientExecute(self, filters=(), filter=''):
        if len(filters) != 0:
            filter = ' OR '.join(filters)
        else:
            filter = self._getNullFilter()
        self._query.setFilter(filter)
        self._recipient.ApplyFilter = False
        self._recipient.Filter = self._query.getFilter()
        self._recipient.ApplyFilter = True
        print("wizardpage._rowRecipientExecute()1 ************** %s" % self._recipient.ActiveCommand)
        self._recipient.execute()
        print("wizardpage._rowRecipientExecute()2 ************** %s" % filter)
        return True

    def _getQuery(self, create, name='smtpMailerOOo'):
        queries = self._database.getQueryDefinitions()
        if queries.hasByName(name):
            query = queries.getByName(name)
        elif create:
            query = self.ctx.ServiceManager.createInstance('com.sun.star.sdb.QueryDefinition')
            queries.insertByName(name, query)
        else:
            query = self._createQuery()
        return query

    def _createQuery(self):
        query = self.ctx.ServiceManager.createInstance('com.sun.star.sdb.QueryDefinition')
        table = self.Connection.getTables().getByIndex(0)
        column = table.getColumns().getByIndex(0)
        query.Command = 'SELECT * FROM "%s" WHERE 0=1 ORDER BY "%s"' % (table.Name, column.Name)
        query.UpdateTableName = table.Name
        return query

    def _getDatabase(self, datasource):
        databases = self.ctx.ServiceManager.createInstance('com.sun.star.sdb.DatabaseContext')
        if databases.hasByName(datasource):
            return databases.getByName(datasource)
        return None

    def saveSetting(self, window):
        if self._modified:
            doc, form = self._getForm(True)
            self._saveTableSetting(doc, window, 'ListBox3', 'PrimaryTable')
            self._emailcolumns = self._saveColumnSetting(doc, window, 'ListBox4', 'EmailColumns')
            self._indexcolumns = self._saveColumnSetting(doc, window, 'ListBox5', 'IndexColumns')
            form.store()
            form.close()
            self._modified = False

    def saveSelection(self):
        if self._modified:
            query = self._getQuery(True)
            command = self._query.getQueryWithSubstitution()
            query.Command = command
            print("wizardpage.saveSelection() %s" % command)
            query.UpdateTableName = self._address.Command
            self._modified = False

    def _saveTableSetting(self, doc, window, id, property):
        table = window.getControl(id).getSelectedItem()
        self._setDocumentUserProperty(doc, property, table)
        print("wizardpage._saveTableSetting() ************** %s" % table)

    def _saveColumnSetting(self, doc, window, id, property):
        items = getStringItemList(window.getControl(id))
        self._setDocumentUserProperty(doc, property, ','.join(items))
        print("wizardpage._saveControlSetting() ************** %s" % ','.join(items))
        return items

    def _getForm(self, create, name='smtpMailerOOo'):
        forms = self._database.DatabaseDocument.getFormDocuments()
        if forms.hasByName(name):
            form = forms.getByName(name)
        elif create:
            form = self._createForm(forms, name)
        else:
            return None, None
        args = getPropertyValueSet({'ActiveConnection': self.Connection,
                                    'OpenMode': 'openDesign',
                                    'Hidden': True})
        doc = forms.loadComponentFromURL(name, '', 0, args)
        return doc, form

    def _createForm(self, forms, name):
        service = 'com.sun.star.sdb.DocumentDefinition'
        args = getPropertyValueSet({'Name': name, 'ActiveConnection': self.Connection})
        form = forms.createInstanceWithArguments(service, args)
        forms.insertByName(name, form)
        form = forms.getByName(name)
        return form

    def _updateControl(self, window, control):
        try:
            tag = control.Model.Tag
            if tag == 'DataSource':
                window.getControl('ListBox1').Model.StringItemList = self.DataSources
            elif tag == 'Columns':
                button = window.getControl('CommandButton2')
                button.Model.Enabled = self._canAdd(control, window.getControl('ListBox4'))
                button = window.getControl('CommandButton6')
                button.Model.Enabled = self._canAdd(control, window.getControl('ListBox5'))
            elif tag == 'EmailAddress':
                indexmax = control.ItemCount -1
                position = control.getSelectedItemPos()
                window.getControl('CommandButton3').Model.Enabled = position != -1
                window.getControl('CommandButton4').Model.Enabled = position > 0
                window.getControl('CommandButton5').Model.Enabled = -1 < position < indexmax
            elif tag == 'PrimaryKey':
                position = control.getSelectedItemPos()
                window.getControl('CommandButton7').Model.Enabled = position != -1
            elif tag == 'Addresses':
                selected = control.hasSelectedRows()
                enabled = control.Model.GridDataModel.RowCount != 0
                window.getControl('CommandButton1').Model.Enabled = enabled
                window.getControl('CommandButton2').Model.Enabled = selected
            elif tag == 'Recipients':
                selected = control.hasSelectedRows()
                enabled = control.Model.GridDataModel.RowCount != 0
                print("WizardHandler._updateControl() %s - %s" % (selected, enabled))
                window.getControl('CommandButton3').Model.Enabled = selected
                window.getControl('CommandButton4').Model.Enabled = enabled
        except Exception as e:
            print("WizardHandler._updateControl() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _canAdd(self, control, listbox):
        enabled = control.getSelectedItemPos() != -1
        if enabled and listbox.ItemCount != 0:
            column = control.getSelectedItem()
            columns = getStringItemList(listbox)
            return column not in columns
        return enabled

    def _getFilter(self, any=False):
        filters = []
        print("WizardHandler._getFilter()1 %s" % (self.ColumnNames, ))
        for column in self._emailcolumns:
            if column in self.ColumnNames:
                print("WizardHandler._getFilter()2 %s" % (column, ))
                filters.append('"%s" IS NOT NULL' % column)
        filter = self._addFilter(filters, any)
        print("WizardHandler._getFilter()3 %s" % filter)
        return filter

    def _getOrderIndex(self):
        index = []
        orders = getRowSetOrders(self._recipient)
        columns = self.ColumnNames
        for order in orders:
            if order in columns:
                index.append(columns.index(order))
        return tuple(index)

    def _getRecipientFilters(self, rows=()):
        filters = []
        if self._recipient.RowCount > 0:
            self._recipient.beforeFirst()
            while self._recipient.next():
                row = self._recipient.Row -1
                if row not in rows:
                    filters.append(self._getFilters(self._recipient))
        print("wizardhandler._getRecipientFilters() %s - %s)" % (self._recipient.RowCount, filters))
        return tuple(filters)

    def _getAddressFilters(self, rows, recipients):
        filters = []
        for row in rows:
            self._address.absolute(row + 1)
            filter = self._getFilters(self._address)
            if filter not in recipients:
                filters.append(filter)
        print("wizardhandler._getAddressFilters() %s - %s)" % (self._address.RowCount, filters))
        return tuple(filters)

    def _getFilters(self, rowset):
        filters = []
        for column in self._indexcolumns:
            if column in self.ColumnNames:
                filter = '"%s"' % column
                i = rowset.findColumn(column)
                filter = "%s = '%s'" % (filter, getValueFromResult(rowset, i))
                filters.append(filter)
        return self._addFilter(filters)

    def _getNullFilter(self):
        filters = []
        for column in self._indexcolumns:
            if column in self.ColumnNames:
                filter = '"%s" IS NULL' % column
                filters.append(filter)
        return self._addFilter(filters)

    def _addFilter(self, filters, any=False):
        separator = ' OR ' if any else ' AND '
        filter = separator.join(filters)
        if len(filters) > 1:
            filter = '(%s)' % filter
        return filter

    def _setDocumentUserProperty(self, document, property, value):
        print("wizardhandler._setDocumentUserProperty() %s - %s" % (property, value))
        properties = document.DocumentProperties.UserDefinedProperties
        if properties.PropertySetInfo.hasPropertyByName(property):
            print("wizardhandler._setDocumentUserProperty() setProperty")
            properties.setPropertyValue(property, value)
        else:
            print("wizardhandler._setDocumentUserProperty() addProperty")
            properties.addProperty(property,
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.MAYBEVOID') +
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.BOUND') +
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.REMOVABLE') +
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.MAYBEDEFAULT'),
            value)

    def _getDocumentUserProperty(self, document, property, default=None):
        print("wizardhandler._getDocumentUserProperty() %s" % (property, ))
        properties = document.DocumentProperties.UserDefinedProperties
        if properties.PropertySetInfo.hasPropertyByName(property):
            print("wizardhandler._getDocumentUserProperty() getProperty")
            return properties.getPropertyValue(property)
        elif default is not None:
            self._setDocumentUserProperty(property, default)
        return default

    def _getPropertySetInfo(self):
        properties = {}
        readonly = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.READONLY')
        transient = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.TRANSIENT')
        properties['Connection'] = getProperty('Connection', 'com.sun.star.sdbc.XConnection', transient)
        properties['DataSources'] = getProperty('DataSources', '[]string', transient)
        properties['TableNames'] = getProperty('TableNames', '[]string', transient)
        properties['ColumnNames'] = getProperty('ColumnNames', '[]string', transient)
        return properties


class CommandEnvironment(unohelper.Base,
                         XCommandEnvironment):

    # XCommandEnvironment
    def getInteractionHandler(self):
        print("CommandEnvironment.getInteractionHandler()")
        return None
    def getProgressHandler(self):
        print("CommandEnvironment.getProgressHandler()")
        return None
