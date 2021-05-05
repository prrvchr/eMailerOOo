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

from com.sun.star.beans import PropertyValue

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.uno import Exception as UnoException

from com.sun.star.document.MacroExecMode import ALWAYS_EXECUTE_NO_WARN

from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.sdb.SQLFilterOperator import EQUAL
from com.sun.star.sdb.SQLFilterOperator import SQLNULL
from com.sun.star.sdb.SQLFilterOperator import NOT_SQLNULL

from smtpserver import createService
from smtpserver import executeDispatch
from smtpserver import getConfiguration
from smtpserver import getDesktop
from smtpserver import getInteractionHandler
from smtpserver import getMessage
from smtpserver import getPathSettings
from smtpserver import getPropertyValue
from smtpserver import getPropertyValueSet
from smtpserver import getSequenceFromResult
from smtpserver import getSimpleFile
from smtpserver import getSqlQuery
from smtpserver import getStringResource
from smtpserver import getUrl
from smtpserver import getValueFromResult
from smtpserver import logMessage

from smtpserver import MailModel

from smtpserver import GridModel
from smtpserver import ColumnModel

from smtpserver import g_identifier
from smtpserver import g_extension
from smtpserver import g_fetchsize

from collections import OrderedDict
from six import string_types
from threading import Thread
from threading import Condition
from time import sleep
import json
import traceback


class MergerModel(MailModel):
    def __init__(self, ctx, datasource):
        self._ctx = ctx
        self._datasource = datasource
        self._configuration = getConfiguration(ctx, g_identifier, True)
        self._document = getDesktop(ctx).CurrentComponent
        #mri = createService(self._ctx, 'mytools.Mri')
        #mri.inspect(self._document)
        self._path = self._getPath()
        self._grid1 = GridModel()
        self._grid2 = GridModel()
        self._column1 = ColumnModel(ctx)
        self._column2 = ColumnModel(ctx)
        self._widths = self._getGridColumnsWidth('MergerGridColumns')
        self._maxcolumns = 8
        self._prefix = 'Sub.'
        service = 'com.sun.star.sdb.DatabaseContext'
        self._dbcontext = createService(ctx, service)
        self._addressbook = None
        self._statement = None
        self._queries = None
        self._address = None
        self._recipient = None
        self._composer = None
        self._subcomposer = None
        self._command = None
        self._subcommand = None
        self._name = None
        self._tables = ()
        self._table = None
        self._query = None
        self._row = 0
        self._changed = False
        self._executed = False
        self._disposed = False
        self._similar = False
        self._temp = False
        self._lock = Condition()
        self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._resources = {'Step': 'MergerPage%s.Step',
                           'Title': 'MergerPage%s.Title',
                           'TabTitle': 'MergerTab%s.Title',
                           'Progress': 'MergerPage1.Label6.Label.%s',
                           'Error': 'MergerPage1.Label8.Label.%s',
                           'Index': 'MergerPage1.Label13.Label.%s',
                           'Message': 'MergerTab2.Label1.Label',
                           'Recipient': 'MailWindow.Label4.Label',
                           'PickerTitle': 'Mail.FilePicker.Title',
                           'Property': 'Mail.Document.Property.%s',
                           'Document': 'MailWindow.Label8.Label.1'}

    @property
    def Connection(self):
        return self._statement.getConnection()

    def _isConnectionNotClosed(self):
        return self._statement is not None

    def _closeConnection(self):
        connection = self._statement.getConnection()
        #self._statement.close()
        self._statement.dispose()
        self._statement = None
        self._address.dispose()
        self._address = None
        self._recipient.dispose()
        self._recipient = None
        self._composer.dispose()
        self._composer = None
        self._subcomposer.dispose()
        self._subcomposer = None
        self._queries.dispose()
        self._queries = None
        self._addressbook.DatabaseDocument.dispose()
        self._addressbook.dispose()
        self._addressbook = None
        connection.close()
        #connection.dispose()

# Procedures called by WizardController
    def dispose(self):
        if self._isConnectionNotClosed():
            self._closeConnection()
        self._datasource.dispose()
        self._disposed = True

# Procedures called by WizardPage1
    # AddressBook methods
    def getAvailableAddressBooks(self):
        return self._dbcontext.getElementNames()

    def getDocumentAddressBook(self):
        addressbook = ''
        service = 'com.sun.star.text.TextDocument'
        if self._document.supportsService(service):
            service = 'com.sun.star.document.Settings'
            settings = self._document.createInstance(service)
            addressbook = settings.CurrentDatabaseDataSource
        return addressbook

    def isAddressBookNotLoaded(self, addressbook):
        return self._addressbook is None or self._addressbook.Name != addressbook

    def setAddressBook(self, *args):
        Thread(target=self._setAddressBook, args=args).start()

    # AddressBook private methods
    def _setAddressBook(self, addressbook, progress, setAddressBook):
        try:
            step = 2
            sleep(0.2)
            progress(5)
            queries = label = msg = None
            self._tables = ()
            self._query = self._table = self._command = self._subcommand = None
            progress(10)
            # TODO: If an addressbook has been loaded we need to close the connection
            # TODO: and dispose all components who use the connection
            if self._isConnectionNotClosed():
                self._closeConnection()
            progress(20)
            try:
                datasource = self._getDataSource(addressbook)
                progress(30)
                if datasource.IsPasswordRequired:
                    handler = getInteractionHandler(self._ctx)
                    connection = datasource.getIsolatedConnectionWithCompletion(handler)
                else:
                    connection = datasource.getIsolatedConnection('', '')
                progress(40)
                service = 'com.sun.star.sdb.SingleSelectQueryComposer'
                if service not in connection.getAvailableServiceNames():
                    msg = self._getErrorMessage(2, service)
                    e = self._getUnoException(msg)
                    raise e
            except UnoException as e:
                format = (addressbook, e.Message)
                msg = self._getErrorMessage(0, format)
            else:
                progress(50)
                self._setConnectionDataSource(connection, datasource)
                self._addressbook = datasource
                self._statement = connection.createStatement()
                self._composer = connection.createInstance(service)
                self._subcomposer = connection.createInstance(service)
                #mri = createService(self._ctx, 'mytools.Mri')
                #mri.inspect(connection)
                progress(60)
                self._similar, self._tables = self._getTablesInfos()
                progress(70)
                self._address = self._getRowSet(connection, TABLE)
                self._recipient = self._getRowSet(connection, COMMAND)
                progress(80)
                self._queries = datasource.getQueryDefinitions()
                progress(90)
                queries = self._getQueries()
                label = self._getIndexLabel()
                step = 3
            progress(100)
            setAddressBook(step, queries, self._tables, label, msg)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def _getDataSource(self, addressbook):
        # We need to check if the registered datasource has an existing odb file
        sf = getSimpleFile(self._ctx)
        location = self._dbcontext.getDatabaseLocation(addressbook)
        if not sf.exists(location):
            msg = self._getErrorMessage(1, location)
            e = self._getUnoException(msg)
            raise e
        if self._temp:
            datasource = self._getTempDataSource(sf, addressbook, location)
        else:
            datasource = self._dbcontext.getByName(addressbook)
        return datasource

    def _getTempDataSource(self, sf, addressbook, location):
        # TODO: We can undo all changes if the wizard is canceled
        # TODO: or abort the Wizard while keeping the work already done
        # TODO: The wizard must be modified to take into account the Cancel button
        url = self._getTempUrl(addressbook)
        if not sf.exists(url):
            sf.copy(location, url)
        datasource = self._dbcontext.getByName(url)
        return datasource

    def _setConnectionDataSource(self, connection, datasource):
        # TODO: It is a workaround to have the Datasource under the Connection
        if not connection.getParent().DatabaseDocument.hasLocation():
            connection.setParent(datasource)

    def _getUnoException(self, msg):
        e = UnoException()
        e.Message = msg
        e.Context = self
        return e

    def _getTempUrl(self, addressbook):
        temp = getPathSettings(self._ctx).Temp
        url = '%s/%s.odb' % (temp, addressbook)
        return url

    def _getTablesInfos(self):
        similar = True
        tables = self.Connection.getTables()
        if tables.hasElements():
            columns = tables.getByIndex(0).getColumns().getElementNames()
            for index in range(1, tables.getCount()):
                table = tables.getByIndex(index)
                if columns != table.getColumns().getElementNames():
                    similar = False
                    break
        return similar, tables.getElementNames()

    def _getQueries(self):
        names = self._queries.getElementNames()
        queries = [name for name in names if self._hasSubQuery(names, name)]
        return tuple(queries)

    def _hasSubQuery(self, names, name):
        subquery = self._getSubQueryName(name)
        return not name.startswith(self._prefix) and subquery in names

    # AddressBook Table methods
    def setAddressBookTable(self, table):
        columns = self._getTableColumns(table)
        return columns

    # Query methods
    def validateQuery(self, query):
        valid = False
        if query != '' and query not in self._tables:
            queries = self._queries.getElementNames()
            valid = not query.startswith(self._prefix) and query not in queries
        return valid

    def setQuery(self, query):
        self._setComposerCommand(query, self._composer)
        self._setRowSet(self._recipient, self._composer)
        name = self._getSubQueryName(query)
        self._setComposerCommand(name, self._subcomposer)
        self._setRowSet(self._address, self._subcomposer)
        table = self._getSubComposerTable()
        return table

    def addQuery(self, table, query):
        name = self._getSubQueryName(query)
        self._addSubQuery(name, table)
        command = getSqlQuery(self._ctx, 'getQueryCommand', table)
        self._addQuery(query, command)
        self._addressbook.DatabaseDocument.store()

    def removeQuery(self, query):
        self._queries.removeByName(query)
        self._removeSubQuery(query)
        self._addressbook.DatabaseDocument.store()

    # Query private shared method
    def _getSubQueryName(self, query):
        return self._prefix + query

    def _setComposerCommand(self, name, composer):
        command = self._queries.getByName(name).Command
        composer.setQuery(command)

    def _addSubQuery(self, name, table):
        command = getSqlQuery(self._ctx, 'getQueryCommand', table)
        query = self._createQuery(command)
        self._queries.insertByName(name, query)

    def _addQuery(self, name, command):
        query = self._createQuery(command)
        if self._similar:
            index = self.getIndex()
            filters = self._getIndexFilters(index)
            self._composer.setQuery(query.Command)
            self._composer.setStructuredFilter(filters)
            query.Command = self._composer.getQuery()
        self._queries.insertByName(name, query)

    def _createQuery(self, command):
        service = 'com.sun.star.sdb.QueryDefinition'
        query = createService(self._ctx, service)
        query.Command = command
        return query

    def _getIndexFilters(self, index):
        filters = ((), )
        if index is not None:
            filter = getPropertyValue(index, 'IS NULL', 0, SQLNULL)
            filters = ((filter, ), )
        return filters

    def _removeSubQuery(self, query):
        name = self._getSubQueryName(query)
        self._queries.removeByName(name)

    def _getSubComposerTable(self):
        return self._subcomposer.getTables().getByIndex(0).Name

    # Email methods
    def getEmails(self):
        emails = []
        filters = self._subcomposer.getStructuredFilter()
        for filter in filters:
            if len(filter) > 0:
                emails.append(filter[0].Name)
        return emails

    def addEmail(self, query, email):
        name = self._getSubQueryName(query)
        emails = self._addEmail(name, email)
        self._addressbook.DatabaseDocument.store()
        return emails

    def removeEmail(self, query, email, table):
        name = self._getSubQueryName(query)
        emails = self._removeEmail(name, email)
        self._addressbook.DatabaseDocument.store()
        enabled = self.canAddColumn(table)
        return emails, enabled

    def moveEmail(self, query, email, position):
        name = self._getSubQueryName(query)
        emails = self._moveEmail(name, email, position)
        self._addressbook.DatabaseDocument.store()
        return emails

    # Email private shared methods
    def _addEmail(self, name, email):
        emails = self.getEmails()
        if email not in emails:
            emails.append(email)
        filters = self._getSubQueryFilters(emails)
        self._subcomposer.setStructuredFilter(filters)
        self._setRowSetFilter(self._address, self._subcomposer)
        self._setQueryCommand(name, self._subcomposer)
        return emails

    def _removeEmail(self, name, email):
        emails = self.getEmails()
        if email in emails:
            emails.remove(email)
        filters = self._getSubQueryFilters(emails)
        self._subcomposer.setStructuredFilter(filters)
        self._setRowSetFilter(self._address, self._subcomposer)
        self._setQueryCommand(name, self._subcomposer)
        return emails

    def _moveEmail(self, name, email, position):
        emails = self.getEmails()
        if email in emails:
            emails.remove(email)
            if 0 <= position <= len(emails):
                emails.insert(position, email)
        filters = self._getSubQueryFilters(emails)
        self._subcomposer.setStructuredFilter(filters)
        self._setRowSetFilter(self._address, self._subcomposer)
        self._setQueryCommand(name, self._subcomposer)
        return emails

    def _getSubQueryFilters(self, emails):
        filters = []
        for email in emails:
            filter = getPropertyValue(email, 'IS NOT NULL', 0, NOT_SQLNULL)
            filters.append((filter, ))
        return tuple(filters)

    # Index methods
    def getIndex(self):
        filters = self._composer.getStructuredFilter()
        index = self._getFiltersIndex(filters)
        return index

    def addIndex(self, query, index):
        filter = getPropertyValue(index, 'IS NULL', 0, SQLNULL)
        filters = (filter, )
        self._setIndex(query, filters)

    def removeIndex(self, query, table):
        filters = ()
        self._setIndex(query, filters)
        enabled = self.canAddColumn(table)
        return enabled

    # Index private shared methods
    def _getFiltersIndex(self, filters):
        index = None
        if len(filters) > 0:
            for filter in filters[0]:
                index = filter.Name
        return index

    def _setIndex(self, query, filter):
        if self._similar:
            self._setQueriesIndex(filter)
        else:
            self._setQueryIndex(query, filter)
        self._addressbook.DatabaseDocument.store()

    def _setQueriesIndex(self, filter):
        for name in self._queries.getElementNames():
            if not name.startswith(self._prefix):
                self._setQueryIndex(name, filter)

    def _setQueryIndex(self, name, filter):
        filters = (filter, )
        self._composer.setStructuredFilter(filters)
        self._setRowSetFilter(self._recipient, self._composer)
        self._setQueryCommand(name, self._composer)

    # Email and Index shared methods
    def canAddColumn(self, table):
        return True if self._similar else table == self._getSubComposerTable()

    # Email and Index private shared methods
    def _setQueryCommand(self, name, composer):
        command = composer.getQuery()
        self._queries.getByName(name).Command = command

    # WizardPage1 commitPage()
    def commitPage1(self, query, updateTravelUI):
        name = self._addressbook.Name
        table = self._getSubComposerTable()
        command = self._composer.getQuery()
        subcommand = self._subcomposer.getQuery()
        loaded = self._name != name
        changed = False if self._similar else self._table != table
        self._changed = loaded or changed
        if self._name is not None:
            # TODO: This will only be executed if WizardPage2 has already been loaded
            if self._changed:
                self._saveColumnWidth(1)
                self._saveColumnWidth(2)

            self._table = table
            if self._subcommand != subcommand:
                if not self._changed:
                    # Update Grid Address only for a change of filters
                    self._updateAddress()
                self._executed = True
                sql = getSqlQuery(self._ctx, 'getRowSetCommand', subcommand)
                self._recipient.Command = sql
                self._updateRecipient(updateTravelUI)
            elif self._command != command:
                self._executed = True
                self._updateRecipient(updateTravelUI)
        else:
            self._table = table
        self._query = query
        self._command = command
        self._subcommand = subcommand

    def _updateAddress(self):
        Thread(target=self._executeAddress).start()

    def _updateRecipient(self, *args):
        Thread(target=self._executeRecipient, args=args).start()

    def _executeAddress(self):
        self._address.execute()

    def _executeRecipient(self, updateTravelUI):
        # TODO: RowSet.RowCount is not accessible during the executuion of RowSet
        # TODO: if self._executed = True then self.getRecipientCount() return 0
        self._recipient.execute()
        self._executed = False
        if self._changed:
            self._initGridColumn(self._recipient, 2)
        updateTravelUI()

# Procedures called by WizardPage2
    def getPageInfos(self, init=False):
        if init:
            self._name = self._addressbook.Name
            self._changed = False
        message = self._getMailingMessage()
        return self._tables, self._table, self._similar, message

    def isChanged(self):
        changed = self._changed
        self._changed = False
        return changed

    def getGridModel(self, tab, width, factor):
        # TODO: com.sun.star.awt.grid.GridColumnModel must be initialized 
        # TODO: before its assignment at com.sun.star.awt.grid.UnoControlGridModel !!!
        if tab == 1:
            grid = self._grid1
            rowset = self._address
        elif tab == 2:
            grid = self._grid2
            rowset = self._recipient
        column = self._getColumnModel(tab)
        widths, titles = self._getGridColumns(tab)
        model = column.getModel(rowset, widths, titles, width, factor)
        return grid, model

    def initGrid(self, address, recipient, initGrid1, initGrid2):
        self._address.addRowSetListener(address)
        self._recipient.addRowSetListener(recipient)
        self._initTab1(initGrid1)
        self._initTab2(initGrid2)

    def _initTab1(self, *args):
        Thread(target=self._initGrid1, args=args).start()

    def _initTab2(self, *args):
        Thread(target=self._initGrid2, args=args).start()

    def _initGrid1(self, initGrid1):
        columns, orders = self._getGridColumnsOrders(self._subcomposer)
        initGrid1(columns, orders)

    def _initGrid2(self, initGrid2):
        command = self._subcomposer.getQuery()
        sql = getSqlQuery(self._ctx, 'getRowSetCommand', command)
        self._recipient.Command = sql
        self._recipient.execute()
        columns, orders = self._getGridColumnsOrders(self._composer)
        message = self.getMailingMessage()
        initGrid2(columns, orders, message)

    def changeAddressRowSet(self):
        self._grid1.setRowSetData(self._address)

    def changeRecipientRowSet(self):
        self._grid2.setRowSetData(self._recipient)

    def setAddressTable(self, table):
        if self._address.Command != table:
            self._address.Command = table
            Thread(target=self._setAddressTable).start()

    def _setAddressTable(self):
        self._address.execute()

    def setAddressColumn(self, *args):
        Thread(target=self._setAddressColumn, args=args).start()

    def setAddressOrder(self, *args):
        Thread(target=self._setAddressOrder, args=args).start()

    def getAddressCount(self):
        return self._address.RowCount

    def setRecipientColumn(self, *args):
        Thread(target=self._setRecipientColumn, args=args).start()

    def setRecipientOrder(self, *args):
        Thread(target=self._setRecipientOrder, args=args).start()

    def getRecipientCount(self):
        print("MergerModel.getRecipientCount() 1 %s" % self._executed)
        if self._executed or self._name != self._addressbook.Name:
            rowcount = 0
        else:
            rowcount = self._recipient.RowCount
        print("MergerModel.getRecipientCount() 2 %s" % rowcount)
        return rowcount

    def addItem(self, *args):
        Thread(target=self._addItem, args=args).start()

    def removeItem(self, *args):
        Thread(target=self._removeItem, args=args).start()

    def setDocumentRecord(self, row):
        if row != self._row:
            self._row = row
            #Thread(target=self._setDocumentRecord).start()

    def getMailingMessage(self):
        message = self._getMailingMessage()
        return message

# Private procedures called by WizardPage2
    def _getGridColumnsOrders(self, composer):
        columns = composer.getColumns().getElementNames()
        orders = composer.getOrderColumns().createEnumeration()
        return columns, orders

    def _getColumnTitles(self, columns):
        titles = OrderedDict()
        for column in columns:
            titles[column] = column
        return titles

    def _getDefaultColumns(self):
        columns = self._getTableColumns(self._table)
        return columns[:self._maxcolumns]

    def _setAddressColumn(self, columns, reset):
        if reset:
            columns = self._getDefaultColumns()
        titles = self._getColumnTitles(columns)
        self._column1.setModel(self._address, titles, reset)

    def _setAddressOrder(self, orders, ascending):
        self._setComposerOrder(self._subcomposer, orders, ascending)
        name = self._getSubQueryName(self._query)
        self._queries.getByName(name).Command = self._subcomposer.getQuery()
        self._addressbook.DatabaseDocument.store()
        self._address.Order = self._subcomposer.getOrder()
        self._address.execute()

    def _setRecipientColumn(self, columns, reset):
        if reset:
            columns = self._getDefaultColumns()
        titles = self._getColumnTitles(columns)
        self._column2.setModel(self._recipient, titles, reset)

    def _setRecipientOrder(self, orders, ascending):
        self._setComposerOrder(self._composer, orders, ascending)
        self._queries.getByName(self._query).Command = self._composer.getQuery()
        self._addressbook.DatabaseDocument.store()
        self._recipient.Order = self._composer.getOrder()
        self._recipient.execute()

    def _setComposerOrder(self, composer, orders, ascending):
        olds, news = self._getComposerOrder(composer, orders)
        composer.setOrder('')
        for order in olds:
            composer.appendOrderByColumn(order, order.IsAscending)
        columns = composer.getColumns()
        for order in news:
            column = columns.getByName(order)
            composer.appendOrderByColumn(column, ascending)

    def _getComposerOrder(self, composer, news):
        olds = []
        enumeration = composer.getOrderColumns().createEnumeration()
        while enumeration.hasMoreElements():
            column = enumeration.nextElement()
            if column.Name in news:
                olds.append(column)
                news.remove(column.Name)
        return olds, news

    def _addItem(self, rows):
        self._updateItem(self._address, rows, True)

    def _removeItem(self, rows):
        self._updateItem(self._recipient, rows, False)

    def _updateItem(self, rowset, rows, add):
        with self._lock:
            filters = self._getRowSetFilters(rowset, rows, add)
            self._composer.setStructuredFilter(filters)
            self._queries.getByName(self._query).Command = self._composer.getQuery()
            self._addressbook.DatabaseDocument.store()
            self._recipient.ApplyFilter = False
            self._recipient.Filter = self._composer.getFilter()
            self._recipient.ApplyFilter = True
            self._recipient.execute()

    def _getRowSetFilters(self, rowset, rows, add):
        index = self.getIndex()
        filters = self._composer.getStructuredFilter()
        values = self._getComposerFilters(filters)
        for row in rows:
            rowset.absolute(row +1)
            self._updateFilters(rowset, index, values, add)
        return self._getStructuredFilters(values)

    def _getComposerFilters(self, values):
        filters = []
        for filter in values:
            filters.append(self._getComposerFilter(filter))
        return filters

    def _getComposerFilter(self, properties):
        filters = []
        for property in properties:
            value = property.Value
            # TODO: Filter can be quoted if its type is String!!!
            if isinstance(value, string_types):
                value = value.strip("'")
            filter = (property.Name, value)
            filters.append(filter)
        return tuple(filters)

    def _updateFilters(self, rowset, index, filters, add):
        filter = self._getRowSetFilter(rowset, index)
        if add:
            if filter not in filters:
                filters.append(filter)
        elif filter in filters:
            filters.remove(filter)

    def _getRowSetFilter(self, rowset, index):
        i = rowset.findColumn(index)
        value = getValueFromResult(rowset, i)
        filter = (index, value)
        filters = (filter, )
        return filters

    def _getStructuredFilters(self, filters):
        structured = []
        for filter in filters:
            properties = []
            for name, value in filter:
                operator = EQUAL if value != 'IS NULL' else SQLNULL
                property = getPropertyValue(name, value, 0, operator)
                properties.append(property)
            structured.append(tuple(properties))
        return tuple(structured)

    def _setDocumentRecord1(self):
        try:
            print("MergerModel._setDocumentRecord() 1")
            url = None
            if self._document.supportsService('com.sun.star.text.TextDocument'):
                url = '.uno:DataSourceBrowser/InsertContent'
            elif self._document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
                url = '.uno:DataSourceBrowser/InsertColumns'
            if url is not None:
                print("MergerModel._setDocumentRecord() 2")
                descriptor = self._getDataDescriptor()
                executeDispatch(self._ctx, url, descriptor)
            print("MergerModel._setDocumentRecord() 3")
        except Exception as e:
            print("MergerModel._setDocumentRecord() ERROR: %s" % traceback.print_exc())

    def _setDocumentRecord(self):
        try:
            print("MergerModel._setDocumentRecord() 1")
            dispatch = None
            frame = self._document.getCurrentController().Frame
            flag = uno.getConstantByName('com.sun.star.frame.FrameSearchFlag.SELF')
            print("MergerModel._setDocumentRecord() 2")
            if self._document.supportsService('com.sun.star.text.TextDocument'):
                url = getUrl(self._ctx, '.uno:DataSourceBrowser/InsertContent')
                dispatch = frame.queryDispatch(url, '_self', flag)
                print("MergerModel._setDocumentRecord() 3")
            elif self._document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
                url = getUrl(self._ctx, '.uno:DataSourceBrowser/InsertColumns')
                dispatch = frame.queryDispatch(url, '_self', flag)
                print("MergerModel._setDocumentRecord() 4")
            if dispatch is not None:
                descriptor = self._getDataDescriptor()
                dispatch.dispatch(url, descriptor)
                print("MergerModel._setDocumentRecord() 5")
        except Exception as e:
            print("MergerModel._setDocumentRecord() ERROR: %s" % traceback.print_exc())

    def _getDataDescriptor(self):
        descriptor = []
        direct = uno.Enum('com.sun.star.beans.PropertyState', 'DIRECT_VALUE')
        descriptor.append(PropertyValue('ResultSet', -1, self._recipient.createResultSet(), direct))
        descriptor.append(PropertyValue('Cursor', -1, self._recipient, direct))
        descriptor.append(PropertyValue('Selection', -1, (self._row, ), direct))
        descriptor.append(PropertyValue('BookmarkSelection', -1, False, direct))
        return tuple(descriptor)


    def _getDataDescriptor1(self):
        properties = {'ResultSet': self._recipient.createResultSet(),
                      'Selection': (self._row, ),
                      'BookmarkSelection': False}
        print("MergerModel._getDataDescriptor() %s" % (properties, ))
        descriptor = getPropertyValueSet(properties)
        return descriptor

# Procedures called by WizardPage3
    def getUrl(self):
        return self._document.URL

    def getDocument(self, url=None):
        if url is None:
            document = self._document
        else:
            properties = {'Hidden': True, 'MacroExecutionMode': ALWAYS_EXECUTE_NO_WARN}
            descriptor = getPropertyValueSet(properties)
            document = getDesktop(self._ctx).loadComponentFromURL(url, '_blank', 0, descriptor)
        return document

    def setUrl(self, url):
        pass

    def initView(self, *args):
        Thread(target=self._initView, args=args).start()

    def getRecipients(self):
        composer = self._getQueryComposer()
        table = composer.getTables().getByIndex(0).Name
        emails, subquery = self._getComposerInfos(table)
        columns = getSqlQuery(self._ctx, 'getComposerColumns', emails)
        filter = composer.getFilter()
        format = (columns, subquery, filter)
        query = getSqlQuery(self._ctx, 'getComposerQuery', format)
        print("MergerModel.getRecipients() %s" % query)
        result = self._statement.executeQuery(query)
        recipients = getSequenceFromResult(result)
        print("MergerModel.getRecipients() %s" % (recipients, ))
        total = len(recipients)
        message = self._getRecipientMessage(total)
        return recipients, message

    def _getComposerInfos(self, table):
        name = self._getComposerName(table)
        composer = self._getTableComposer(name, table)
        emails = self._getEmails(composer)
        subquery =  self._getComposerSubQuery(table, composer)
        return emails, subquery

# Private procedures called by WizardPage3
    def _initView(self, handler, initView, initRecipient):
        self._address.addRowSetListener(handler)
        self._recipient.addRowSetListener(handler)
        initView(self._document)
        recipients, message = self.getRecipients()
        initRecipient(recipients, message)

    def _getDocumentName(self):
        url = None
        location = self._document.getLocation()
        if location:
            url = getUrl(self._ctx, location)
        return None if url is None else url.Name

# MergerModel StringRessoure methods
    def getPageStep(self, pageid):
        resource = self._resources.get('Step') % pageid
        step = self._resolver.resolveString(resource)
        return step

    def getPageTitle(self, pageid):
        resource = self._resources.get('Title') % pageid
        title = self._resolver.resolveString(resource)
        return title

    def getTabTitle(self, tab):
        resource = self._resources.get('TabTitle') % tab
        return self._resolver.resolveString(resource)

    def getProgressMessage(self, value):
        resource = self._resources.get('Progress') % value
        return self._resolver.resolveString(resource)

    def _getErrorMessage(self, code, format):
        resource = self._resources.get('Error') % code
        return self._resolver.resolveString(resource) % format

    def _getIndexLabel(self):
        resource = self._resources.get('Index') % int(self._similar)
        return self._resolver.resolveString(resource)

    def _getRecipientMessage(self, total):
        resource = self._resources.get('Recipient')
        return self._resolver.resolveString(resource) % total

    def _getMailingMessage(self):
        resource = self._resources.get('Message')
        return self._resolver.resolveString(resource) % (self._query, self._table)

# Procedures called internally
    def _getPath(self):
        location = self.getUrl()
        if location != '':
            url = getUrl(self._ctx, location)
            path = url.Protocol + url.Path
        else:
            path = getPathSettings(self._ctx).Work
        return path

    # RowSet private methods
    def _getRowSet(self, connection, command):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.ActiveConnection = connection
        rowset.CommandType = command
        rowset.FetchSize = g_fetchsize
        print("MergerModel._getRowSet() FetchSize = %s" % rowset.FetchSize)
        return rowset

    def _setRowSet(self, rowset, composer):
        self._setRowSetFilter(rowset, composer)
        self._setRowSetOrder(rowset, composer)

    def _setRowSetFilter(self, rowset, composer):
        rowset.ApplyFilter = False
        rowset.Filter = composer.getFilter()
        rowset.ApplyFilter = True

    def _setRowSetOrder(self, rowset, composer):
        rowset.Order = composer.getOrder()

    def _getComposerSubQuery(self, table, composer):
        filter = composer.getFilter()
        format = (table, filter)
        command = getSqlQuery(self._ctx, 'getComposerSubQuery', format)
        return command

    # Table private methods
    def _getTableColumns(self, table):
        table = self.Connection.getTables().getByName(table)
        columns = table.getColumns().getElementNames()
        return columns

    # Composer private methods
    def _getSubQueryOrder(self):
        name = self._getSubQueryName(self._query)
        query = self._queries.getByName(name)
        self._composer.setQuery(query.Command)
        return self._composer.getOrder()

    # Grid Columns private methods
    def _getGridColumnsWidth(self, name):
        config = self._configuration.getByName(name)
        widths = json.loads(config, object_pairs_hook=OrderedDict)
        return widths

    def _initGridColumn(self, rowset, tab):
        # TODO: We do not reset the grid columns if ColumnModel is not yet assigned
        model = self._getColumnModel(tab)
        widths, titles = self._getGridColumns(tab)
        model.initModel(rowset, widths, titles)

    def _getGridColumns(self, tab):
        names = self._widths.get(self._name, {})
        widths = names.get(self._getColumnsName(tab), {})
        if widths:
            titles = self._getColumnTitles(widths)
        else:
            columns = self._getDefaultColumns()
            titles = self._getColumnTitles(columns)
        return widths, titles

    def _saveColumnWidth(self, tab):
        if self._name not in self._widths:
            self._widths[self._name] = {}
        addressbook = self._widths[self._name]
        name = self._getColumnsName(tab)
        addressbook[name] = self._getColumnModel(tab).getWidths()
        print("MergerModel._saveColumnWidth() %s - %s" % (column, self._widths))
        model.resetModel()

    def _getColumnModel(self, tab):
        return self._column1 if tab == 1 else self._column2

    def _getColumnsName(self, tab):
        if self._similar:
            name = 'Grid%s' % tab
        else:
            name = 'Grid%s.%s' % (tab, self._table)
