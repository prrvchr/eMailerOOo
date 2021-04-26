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

from com.sun.star.document.MacroExecMode import ALWAYS_EXECUTE_NO_WARN

from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.sdb.SQLFilterOperator import EQUAL
from com.sun.star.sdb.SQLFilterOperator import SQLNULL
from com.sun.star.sdb.SQLFilterOperator import NOT_SQLNULL

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

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
        service = 'com.sun.star.sdb.DatabaseContext'
        self._dbcontext = createService(ctx, service)
        self._address = self._getRowSet(TABLE, g_fetchsize)
        self._recipient = self._getRowSet(QUERY, g_fetchsize)
        self._grid1 = GridModel()
        self._grid2 = GridModel()
        self._column1 = ColumnModel(ctx)
        self._column2 = ColumnModel(ctx)
        self._width1 = self._getColumnsWidth('MergerGrid1Columns')
        self._width2 = self._getColumnsWidth('MergerGrid2Columns')
        self._maxcolumns = 8
        self._composers = {}
        self._composer = None
        self._queries = None
        self._addressbook = None
        self._database = None
        self._tables = ()
        self._table = None
        self._query = None
        self._previous = None
        self._statement = None
        self._row = 0
        self._loaded = False
        self._disposed = False
        self._similar = False
        self._filtered = False
        self._changed = False
        self._updated = False
        self._lock = Condition()
        self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._resources = {'Step': 'MergerPage%s.Step',
                           'Title': 'MergerPage%s.Title',
                           'TabTitle': 'MergerTab%s.Title',
                           'TabTip': 'MergerTab%s.Tab.ToolTip',
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

    def setAddressBook(self, *args):
        Thread(target=self._setAddressBook, args=args).start()

    # AddressBook private methods
    def _setAddressBook(self, addressbook, progress, setAddressBook):
        sleep(0.2)
        progress(5)
        step = 2
        queries = label = msg = None
        progress(10)
        # TODO: We do not save the grid columns for the first load of a new self._addressbook
        if self._addressbook is not None:
            if self._table is not None:
                self._saveColumnWidth(self._column1, self._width1, self._table)
            if self._query is not None:
                self._saveColumnWidth(self._column2, self._width2, self._query)
        self._previous = self._query = self._table = None
        self._filtered = self._changed = self._updated = False
        self._addressbook = addressbook
        progress(20)
        try:
            database = self._getDatabase()
            progress(30)
            if database.IsPasswordRequired:
                handler = getInteractionHandler(self._ctx)
                connection = database.connectWithCompletion(handler)
            else:
                connection = database.getConnection('', '')
            progress(40)
            service = 'com.sun.star.sdb.SingleSelectQueryComposer'
            if service not in connection.getAvailableServiceNames():
                msg = self._getErrorMessage(2, service)
                e = self._getUnoException(msg)
                raise e
        except UnoException as e:
            self._addressbook = None
            format = (addressbook, e.Message)
            msg = self._getErrorMessage(0, format)
        else:
            progress(50)
            self._database = database
            self._statement = connection.createStatement()
            self._composer = self.Connection.createInstance(service)
            progress(60)
            self._setTablesInfos()
            progress(70)
            #self._address.ActiveConnection = connection
            #self._recipient.ActiveConnection = connection
            url = self._getDataSourceTempUrl()
            self._address.DataSourceName = url
            self._recipient.DataSourceName = url
            #mri = createService(self._ctx, 'mytools.Mri')
            #mri.inspect(connection)
            #mri.inspect(self._recipient)
            progress(80)
            self._queries = database.getQueryDefinitions()
            progress(90)
            queries = self._getQueries()
            label = self._getIndexLabel()
            step = 3
        progress(100)
        setAddressBook(step, queries, self._tables, label, msg)

    def _getDatabase1(self, addressbook):
        database = None
        if self._dbcontext.hasByName(addressbook):
            database = self._dbcontext.getByName(addressbook)
        return database

    def _getDatabase(self):
        sf = getSimpleFile(self._ctx)
        location = self._dbcontext.getDatabaseLocation(self._addressbook)
        if not sf.exists(location):
            msg = self._getErrorMessage(1, location)
            e = self._getUnoException(msg)
            raise e
        url = self._getDataSourceTempUrl()
        if not sf.exists(url):
            sf.copy(location, url)
        database = self._dbcontext.getByName(url)
        return database

    def _getUnoException(self, msg):
        e = UnoException()
        e.Message = msg
        e.Context = self
        return e

    def _getDataSourceTempUrl(self):
        temp = getPathSettings(self._ctx).Temp
        url = '%s/%s.odb' % (temp, self._addressbook)
        return url

    def _setTablesInfos(self):
        self._similar = True
        tables = self.Connection.getTables()
        self._tables = tables.getElementNames()
        if len(self._tables) > 0:
            columns = tables.getByIndex(0).getColumns().getElementNames()
            for index in range(1, tables.getCount()):
                table = tables.getByIndex(index)
                if columns != table.getColumns().getElementNames():
                    self._similar = False
                    break

    def _getQueries(self):
        prefix = self._getSubQueryPrefix()
        names = self._queries.getElementNames()
        queries = [name for name in names if not name.startswith(prefix)]
        return tuple(queries)

    # AddressBook Table methods
    def setAddressBookTable(self, table):
        columns = self._getTableColumns(table)
        return columns

    # Query methods
    def validateQuery(self, query, exist):
        valid = False
        if not exist and query != '':
            queries = self._queries.getElementNames()
            valid = all((query not in self._tables,
                         query not in queries))
        return valid

        # Query addQuery() method begin
    def addQuery(self, table, query):
        name = self._getSubQueryName(query)
        self._addSubQuery(name, table)
        command = getSqlQuery(self._ctx, 'getQueryCommand', name)
        self._addQuery(query, command)
        self._database.DatabaseDocument.store()

    def _addSubQuery(self, name, table):
        command = getSqlQuery(self._ctx, 'getQueryCommand', table)
        query = self._createQuery(command)
        self._queries.insertByName(name, query)

    def _addQuery(self, name, command):
        query = self._createQuery(command)
        if self._similar:
            index = self._getQueriesIndex()
            filters = self._getIndexFilters(index)
            self._setQueryFilters(query, filters)
        self._queries.insertByName(name, query)

    def _getQueriesIndex(self):
        index = None
        prefix = self._getSubQueryPrefix()
        for name in self._queries.getElementNames():
            if not name.startswith(prefix):
                filters = self._getQueryFilters(name)
                index = self._getFiltersIndex(filters)
                break
        return index

    def _getQueryFilters(self, name):
        query = self._queries.getByName(name)
        self._composer.setQuery(query.Command)
        filters = self._composer.getStructuredFilter()
        return filters

    def _getFiltersIndex(self, filters):
        index = None
        if len(filters) > 0:
            for filter in filters[0]:
                index = filter.Name
        return index

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

    def _setQueryFilters(self, query, filters):
        self._composer.setQuery(query.Command)
        self._composer.setStructuredFilter(filters)
        query.Command = self._composer.getQuery()
        # Query addQuery() method end

        # Query removeQuery() method begin
    def removeQuery(self, query):
        self._queries.removeByName(query)
        self._removeSubQuery(query)
        self._database.DatabaseDocument.store()

    def _removeSubQuery(self, query):
        name = self._getSubQueryName(query)
        self._queries.removeByName(name)
        # Query removeQuery() method end

    # Query private shared methods
    def _getSubQueryPrefix(self):
        return '%s.' % self._addressbook

    def _getSubQueryName(self, query):
        return self._getSubQueryPrefix() + query

    # Email methods
        # Email getEmails() method begin
    def getEmails(self, query, exist):
        emails = ()
        if exist:
            name = self._getSubQueryName(query)
            emails = self._getEmails(name)
        return emails

    def _getEmails(self, name):
        query = self._queries.getByName(name)
        emails = self._getQueryEmails(query, True)
        return emails
        # Email getEmails() method end

        # Email addEmail() method begin
    def addEmail(self, query, email):
        name = self._getSubQueryName(query)
        emails = self._addEmail(name, email)
        self._database.DatabaseDocument.store()
        self._setFiltered()
        return emails

    def _addEmail(self, name, email):
        query = self._queries.getByName(name)
        emails = self._getQueryEmails(query)
        if email not in emails:
            emails.append(email)
        filters = self._getSubQueryFilters(emails)
        self._composer.setStructuredFilter(filters)
        query.Command = self._composer.getQuery()
        self._address.Filter = self._composer.getFilter()
        return emails
        # Email addEmail() method end

        # Email removeEmail() method begin
    def removeEmail(self, query, email):
        name = self._getSubQueryName(query)
        emails = self._removeEmail(name, email)
        self._database.DatabaseDocument.store()
        self._setFiltered()
        return emails

    def _removeEmail(self, name, email):
        query = self._queries.getByName(name)
        emails = self._getQueryEmails(query)
        if email in emails:
            emails.remove(email)
        filters = self._getSubQueryFilters(emails)
        self._composer.setStructuredFilter(filters)
        query.Command = self._composer.getQuery()
        self._address.Filter = self._composer.getFilter()
        return emails
        # Email removeEmail() method end

        # Email moveEmail() method begin
    def moveEmail(self, query, email, index):
        name = self._getSubQueryName(query)
        emails = self._moveEmail(name, email, index)
        self._database.DatabaseDocument.store()
        self._setFiltered()
        return emails

    def _moveEmail(self, name, email, index):
        query = self._queries.getByName(name)
        emails = self._getQueryEmails(query)
        if email in emails:
            emails.remove(email)
            if 0 <= index <= len(emails):
                emails.insert(index, email)
        filters = self._getSubQueryFilters(emails)
        self._composer.setStructuredFilter(filters)
        query.Command = self._composer.getQuery()
        self._address.Filter = self._composer.getFilter()
        return emails
        # Email moveEmail() method end

    # Email private shared methods
    def _getQueryEmails(self, query, update=False):
        self._composer.setQuery(query.Command)
        filters = self._composer.getStructuredFilter()
        if update:
            self._address.Filter = self._composer.getFilter()
        emails = self._getSubQueryEmails(filters)
        return emails

    def _getSubQueryEmails(self, filters):
        emails = []
        for filter in filters:
            if len(filter) > 0:
                emails.append(filter[0].Name)
        return emails

    def _getSubQueryFilters(self, emails):
        filters = []
        for email in emails:
            filter = getPropertyValue(email, 'IS NOT NULL', 0, NOT_SQLNULL)
            filters.append((filter, ))
        return tuple(filters)

    def _setFiltered(self):
        self._filtered = self._updated = True

    # Index methods
    def getIndex(self, query, exist):
        index = None
        if exist:
            filters = self._getQueryFilters(query)
            index = self._getFiltersIndex(filters)
        return index

        # Index setIndex() method begin
    def setIndex(self, query, index):
        if self._similar:
            self._setQueriesIndex(index)
        else:
            self._setQueryIndex(query, index)
        self._database.DatabaseDocument.store()
        self._updated = True

    def _setQueriesIndex(self, index):
        prefix = self._getSubQueryPrefix()
        for name in self._queries.getElementNames():
            if not name.startswith(prefix):
                self._setQueryIndex(name, index)

    def _setQueryIndex(self, name, index):
        query = self._queries.getByName(name)
        self._composer.setQuery(query.Command)
        if index is None:
            filters = ((), )
        else:
            filter = getPropertyValue(index, 'IS NULL', 0, SQLNULL)
            filters = ((filter, ), )
        self._composer.setStructuredFilter(filters)
        query.Command = self._composer.getQuery()
        # Index setIndex() method end

    # WizardPage1 commitPage()
    def setQuery(self, query, updateTravelUI):
        if query != self._query:
            self._setChanged()
            self._recipient.Command = query
            self._loaded = False
            self._updateRecipient(updateTravelUI)
        self._previous = self._query
        self._query = query
        self._table = self._getSubQueryTable()
        #if self.isFiltered():
        #    self._updateAddress()
        #if self._isUpdated():


    def _getSubQueryTable(self):
        name = self._getSubQueryName(self._query)
        query = self._queries.getByName(name)
        print("MergerModel._getSubQueryTable() %s" % query.Command)
        self._composer.setQuery(query.Command)
        table = self._composer.getTables().getByIndex(0).Name
        return table

    def _updateAddress(self):
        name = self._getSubQueryName(self._query)
        query = self._queries.getByName(name)
        self._composer.setQuery(query.Command)
        self._address.Filter = self._composer.getFilter()
        #self._address.Order = self._composer.getOrder()
        self._address.ApplyFilter = True

    def _updateRecipient(self, *args):
        Thread(target=self._executeRecipient, args=args).start()

    def _executeRecipient(self, updateTravelUI):
        # TODO: RowSet.RowCount is not accessible during the executuion of RowSet
        # TODO: if self._loaded = False then self.getRecipientCount() return 0
        with self._lock:
            self._recipient.execute()
            self._loaded = True
        updateTravelUI()

    def _setChanged(self):
        self._changed = self._updated = True

    def _isUpdated(self):
        updated = self._updated
        self._updated = False
        return updated

# Procedures called by WizardPage2
    def getFilteredTables(self):
        if self._similar:
            tables = self._tables
        else:
            tables = (self._table, )
        return tables

    def initRowSet(self, address, recipient, initTab1, initTab2):
        self._filtered = self._changed = False
        self._address.addRowSetListener(address)
        self._recipient.addRowSetListener(recipient)
        self._address.Order = self._getSubQueryOrder()
        #self.initTab1(initTab1)
        self.initTab2(initTab2)

    def initTab1(self, *args):
        Thread(target=self._initTab1, args=args).start()

    def _initTab1(self, initTab1):
        with self._lock:
            #self._grid1.setRowSetData(self._address)
            #self._initColumn1()
            name = self._getSubQueryName(self._query)
            query = self._queries.getByName(name)
            self._composer.setQuery(query.Command)
            columns, orders = self._getGridColumnsOrders()
        initTab1(columns, orders)

    def initTab2(self, *args):
        Thread(target=self._initTab2, args=args).start()

    def _initTab2(self, initTab2):
        with self._lock:
            self._grid2.setRowSetData(self._recipient)
            self._initColumn2()
            query = self._queries.getByName(self._query)
            self._composer.setQuery(query.Command)
            columns, orders = self._getGridColumnsOrders()
        message = self.getMailingMessage()
        initTab2(columns, orders, message)

    def isFiltered(self):
        filtered = self._filtered
        self._filtered = False
        return filtered

    def isChanged(self):
        changed = self._changed
        self._changed = False
        return changed

    def getGridModels(self, tab, width, factor, tables=None):
        # TODO: com.sun.star.awt.grid.GridColumnModel must be initialized 
        # TODO: before its assignment at com.sun.star.awt.grid.UnoControlGridModel !!!
        if tab == 1:
            model = self._grid1
            widths, titles = self._getGridColumns(self._width1, self._table)
            column = self._column1.getColumnModel(self._address, widths, titles, width, factor)
        elif tab == 2:
            model = self._grid2
            widths, titles = self._getGridColumns(self._width2, self._query)
            column = self._column2.getColumnModel(self._recipient, widths, titles, width, factor)
        return model, column

    def addressChanged(self):
        self._grid1.setRowSetData(self._address)

    def recipientChanged(self):
        self._grid2.setRowSetData(self._recipient)

    def updateColumn2(self, *args):
        if not self._similar:
            Thread(target=self._updateColumn2, args=args).start()

    def _updateColumn2(self, updateColumn2):
        with self._lock:
            if self._previous is not None:
                self._saveColumnWidth(self._column2, self._width2, self._previous)
            self._initColumn2()
            query = self._queries.getByName(self._query)
            self._composer.setQuery(query.Command)
            columns, orders = self._getGridColumnsOrders()
            updateColumn2(columns, orders)

    def setAddressTable(self, table):
        # TODO: We do not save the grid columns for the first change of self._table
        #if self._table is not None:
        #    self._saveColumnWidth(self._column1, self._width1, self._table)
        command = getSqlQuery(self._ctx, 'getQueryCommand', table)
        if self._address.Command != table:
            self._address.Command = table
            self._updateAddress()
        #self._initColumn1()
        #name = self._getSubQueryName(self._query)
        #query = self._queries.getByName(name)
        #columns, orders = self._getGridColumnsOrders(query)
        #return columns, orders

    def _updateAddress(self):
        Thread(target=self._executeAddress).start()

    def _executeAddress(self):
        self._address.ApplyFilter = False
        self._address.ApplyFilter = True
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
        print("MergerModel.getRecipientCount() 1 %s" % self._loaded)
        rowcount = self._recipient.RowCount if self._loaded else 0
        print("MergerModel.getRecipientCount() 2 %s" % rowcount)
        return rowcount

    def addItem(self, *args):
        Thread(target=self._addItem, args=args).start()

    def removeItem(self, *args):
        Thread(target=self._removeItem, args=args).start()

    def setDocumentRecord(self, row):
        if row != self._row:
            self._row = row
            Thread(target=self._setDocumentRecord).start()

    def getMailingMessage(self):
        message = self._getMailingMessage()
        return message

# Private procedures called by WizardPage2
    def _getColumnsWidth(self, name):
        config = self._configuration.getByName(name)
        width = json.loads(config, object_pairs_hook=OrderedDict)
        return width

    def _getGridColumnsOrders(self):
        columns = self._composer.getColumns().getElementNames()
        orders = self._composer.getOrderColumns().createEnumeration()
        return columns, orders

    def _getGridColumns(self, config, name):
        names = config.get(self._addressbook, {})
        widths = names.get(name, {})
        if widths:
            titles = self._getColumnTitles(widths)
        else:
            columns = self._getDefaultColumns()
            titles = self._getColumnTitles(columns)
        return widths, titles

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
        self._column1.setColumnModel(self._address, titles, reset)

    def _setAddressOrder(self, orders, ascending):
        name = self._getSubQueryName(self._query)
        query = self._queries.getByName(name)
        self._composer.setQuery(query.Command)
        self._setGridOrder(orders, ascending)
        command = composer.getQuery()
        query.Command = command
        self._database.DatabaseDocument.store()
        self._address.Order = self._composer.getOrder()
        self._executeAddress()

    def _setRecipientColumn(self, columns, reset):
        if reset:
            columns = self._getDefaultColumns()
        titles = self._getColumnTitles(columns)
        self._column2.setColumnModel(self._recipient, titles, reset)

    def _setRecipientOrder(self, orders, ascending):
        query = self._queries.getByName(self._query)
        self._composer.setQuery(query.Command)
        self._setGridOrder(orders, ascending)
        query.Command = self._composer.getQuery()
        self._database.DatabaseDocument.store()
        self._recipient.execute()

    def _setGridOrder(self, orders, ascending):
        olds, news = self._getGridOrder(orders)
        self._composer.Order = ''
        for order in olds:
            self._composer.appendOrderByColumn(order, order.IsAscending)
        columns = self._composer.getColumns()
        for order in news:
            column = columns.getByName(order)
            self._composer.appendOrderByColumn(column, ascending)

    def _getGridOrder(self, news):
        olds = []
        enumeration = self._composer.getOrderColumns().createEnumeration()
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
            query = self._queries.getByName(self._query)
            self._composer.setQuery(query.Command)
            filters = self._getRowSetFilters(rowset, rows, add)
            self._composer.setStructuredFilter(filters)
            query.Command = self._composer.getQuery()
            self._database.DatabaseDocument.store()
            self._recipient.execute()

    def _getRowSetFilters(self, rowset, rows, add):
        filters = self._composer.getStructuredFilter()
        index = self._getFiltersIndex(filters)
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
        url = self._getDataSourceTempUrl()
        properties = {'DataSourceName': url,
                      'Command': self._recipient.Command,
                      'CommandType': self._recipient.CommandType,
                      'Cursor': self._recipient,
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

    def getTabTip(self, tab):
        resource = self._resources.get('TabTip') % tab
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
    def _getRowSet(self, command, fetchsize):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.CommandType = command
        rowset.FetchSize = fetchsize
        print("MergerModel._getRowSet() FetchSize = %s" % rowset.FetchSize)
        return rowset

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
    def _initColumn1(self):
        # TODO: We do not reset the grid columns if ColumnModel is not yet assigned
        if self._column1.isInitialized():
            widths, titles = self._getGridColumns(self._width1, self._table)
            self._column1.initColumnModel(self._address, widths, titles)

    def _initColumn2(self):
        # TODO: We do not reset the grid columns if ColumnModel is not yet assigned
        if self._column2.isInitialized():
            widths, titles = self._getGridColumns(self._width2, self._query)
            self._column2.initColumnModel(self._recipient, widths, titles)

    def _saveColumnWidth(self, column, widths, name):
        # TODO: We do not save the grid columns if ColumnModel is not yet assigned
        if column.isInitialized():
            columns = column.getColumnWidth()
            if self._addressbook not in widths:
                widths[self._addressbook] = {}
            addressbook = widths[self._addressbook]
            addressbook[name] = columns
            print("MergerModel._saveColumnWidth() %s - %s" % (name, widths))
