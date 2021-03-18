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

from com.sun.star.sdbc import SQLException

from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.sdb.SQLFilterOperator import EQUAL
from com.sun.star.sdb.SQLFilterOperator import SQLNULL
from com.sun.star.sdb.SQLFilterOperator import NOT_SQLNULL

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService
from unolib import getDesktop
from unolib import getStringResource
from unolib import getConfiguration
from unolib import getPropertyValueSet
from unolib import getPropertyValue
from unolib import getUrl

from smtpserver.grid import GridModel
from smtpserver.grid import ColumnModel

from smtpserver.dbtools import getValueFromResult

from smtpserver import g_identifier
from smtpserver import g_extension
from smtpserver import g_fetchsize

from smtpserver import logMessage
from smtpserver import getMessage

from collections import OrderedDict
from six import string_types
from threading import Condition
from threading import Thread
import json
from time import sleep
import traceback


class MergerModel(unohelper.Base):
    def __init__(self, ctx):
        self._ctx = ctx
        self._stringResource = getStringResource(ctx, g_identifier, g_extension)
        self._configuration = getConfiguration(ctx, g_identifier, True)
        self._doc = getDesktop(ctx).CurrentComponent
        service = 'com.sun.star.sdb.DatabaseContext'
        self._dbcontext = createService(ctx, service)
        self._address = self._getRowSet(COMMAND, g_fetchsize)
        self._recipient = self._getRowSet(COMMAND, g_fetchsize)
        self._column1 = ColumnModel(ctx)
        self._column2 = ColumnModel(ctx)
        self._maxcolumns = 8
        self._composers = {}
        self._table = None
        self._tables = ()
        self._query = None
        self._addressbook = None
        self._datasource = None
        self._statement = None
        self._row = 0
        self._similar = False
        self._lock = Condition()

    @property
    def Connection(self):
        return self._statement.getConnection()

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

# Procedures called by WizardPage1
    # DataSource methods
    def getAvailableDataSources(self):
        return self._dbcontext.getElementNames()

    def getDocumentDataSource(self):
        datasource = ''
        if self._doc.supportsService('com.sun.star.text.TextDocument'):
            service = 'com.sun.star.document.Settings'
            datasource = self._doc.createInstance(service).CurrentDatabaseDataSource
        return datasource

    def getDataSource(self):
        return self.Connection.getParent()

    def setDataSource(self, *args):
        Thread(target=self._setDataSource, args=args).start()

    # Table methods
    def getTableColumns(self, name):
        table = self.Connection.getTables().getByName(name)
        columns = table.getColumns().getElementNames()
        return columns

    def setTable(self, table):
        self._table = table
        columns = self._getTableColumns(table)
        composer = self._getTableComposer(table)
        emails = self._getEmails(composer)
        return columns, emails

    # Query methods
    def validateQuery(self, query, exist):
        valid = False
        if not exist and query != '':
            valid = all((query not in self._tables,
                         query not in self._getComposerNames()))
        return valid

    def setQuery(self, query):
        self._query = query
        composer = self._getQueryComposer()
        indexes = self._getIndexes(composer)
        return indexes

    def addQuery(self, table, query):
        composers = self._getComposers()
        if query not in composers:
            indexes = self._getComposerIndexes(composers)
            composer = self._createComposer()
            command = self._getComposerCommand(table)
            composer.setQuery(command)
            if indexes is not None:
                self._setIndexes(composer, indexes)
            composers[query] = composer

    def removeQuery(self, query):
        composers = self._getComposers()
        if query in composers:
            del composers[query]

    # Email methods
    def addEmail(self, table, email):
        composer = self._getTableComposer(table)
        emails = self._getEmails(composer)
        print("MergerModel.addEmail() %s - %s" % (table, emails))
        if email not in emails:
            emails.append(email)
        if self._similar:
            self._setComposerEmails(emails)
        else:
            self._setEmails(composer, emails)
        return emails

    def removeEmail(self, table, email):
        composer = self._getTableComposer(table)
        emails = self._getEmails(composer)
        print("MergerModel.removeEmail() %s - %s" % (table, emails))
        if email in emails:
            emails.remove(email)
        if self._similar:
            self._setComposerEmails(emails)
        else:
            self._setEmails(composer, emails)
        return emails

    def moveEmail(self, table, email, index):
        composer = self._getTableComposer(table)
        emails = self._getEmails(composer)
        print("MergerModel.moveEmail() %s - %s" % (table, emails))
        if email in emails:
            emails.remove(email)
            if 0 <= index <= len(emails):
                emails.insert(index, email)
        if self._similar:
            self._setComposerEmails(emails)
        else:
            self._setEmails(composer, emails)
        return emails

    # Index methods
    def addIndex(self, index):
        composer = self._getQueryComposer()
        indexes = self._getIndexes(composer)
        if index not in indexes:
            indexes.append(index)
        self._setIndexes(composer, indexes)
        if self._similar:
            self._setComposerIndexes(indexes)
        return indexes

    def removeIndex(self, index):
        composer = self._getQueryComposer()
        indexes = self._getIndexes(composer)
        if index in indexes:
            indexes.remove(index)
        self._setIndexes(composer, indexes)
        if self._similar:
            self._setComposerIndexes(indexes)
        return indexes

# Procedures called by WizardPage2
    def getRowSetAddress(self):
        return self._address

    def getRowSetRecipient(self):
        return self._recipient

    def getTables(self):
        tables = self.Connection.getTables().getElementNames()
        if not self._similar:
            names = []
            composers = self._getComposers()
            for table in tables:
                name = self._getComposerName(table)
                if name in composers:
                    composer = composers[name]
                    filter = composer.getFilter()
                    print("MergerModel.getTables() '%s'" % filter)
                    if filter != '':
                        names.append(table)
            tables = tuple(names)
        return tables

    def getDefaultTable(self):
        return self._table

    def addRowSetListener(self, address, recipient):
        self._address.addRowSetListener(address)
        self._recipient.addRowSetListener(recipient)

    def getGridModels(self, tab, width, factor, table=None):
        # TODO: com.sun.star.awt.grid.GridColumnModel must be initialized 
        # TODO: before its assignment at com.sun.star.awt.grid.UnoControlGridModel !!!
        if tab == 1:
            data = GridModel(self._address)
            widths = self._getAddressColumnsWidth(table)
            titles = self._getAddressColumnTitles(widths, table)
            self._column1.initColumnModel(self._address, widths, titles, width, factor)
            column = self._column1.getColumnModel()
        elif tab == 2:
            data = GridModel(self._recipient)
            widths = self._getRecipientColumnsWidth()
            titles = self._getRecipientColumnTitles(widths)
            self._column2.initColumnModel(self._recipient, widths, titles, width, factor)
            column = self._column2.getColumnModel()
        return data, column

    def setAddressBook(self, *args):
        Thread(target=self._setAddressBook, args=args).start()

    def setAddressColumn(self, *args):
        Thread(target=self._setAddressColumn, args=args).start()

    def setRecipientColumn(self, *args):
        Thread(target=self._setRecipientColumn, args=args).start()

    def setAddressOrder(self, *args):
        Thread(target=self._setAddressOrder, args=args).start()

    def setRecipientOrder(self, *args):
        Thread(target=self._setRecipientOrder, args=args).start()

    def addItem(self, *args):
        Thread(target=self._addItem, args=args).start()

    def removeItem(self, *args):
        Thread(target=self._removeItem, args=args).start()

    def setDocumentRecord(self, row):
        if row != self._row:
            self._row = row
            Thread(target=self._setDocumentRecord).start()

    def getAddressCount(self):
        return self._address.RowCount

    def getRecipientCount(self):
        return self._recipient.RowCount

# Private procedures called by WizardPage2
    def _initGrids(self, address, recipient, initGrids):
        self._address.addRowSetListener(address)
        self._recipient.addRowSetListener(recipient)
        with self._lock:
            initGrids(table)

    def _getAddressColumnTitles(self, widths, table):
        if widths:
            titles = self._getColumnTitles(widths)
        else:
            columns = self._getAddressColumns(table)
            titles = self._getColumnTitles(columns)
        return titles

    def _getRecipientColumnTitles(self, widths):
        if widths:
            titles = self._getColumnTitles(widths)
        else:
            columns = self._getRecipientColumns()
            titles = self._getColumnTitles(columns)
        return titles

    def _getAddressColumnsWidth(self, table):
        columns = self._configuration.getByName('MergerGrid1Columns')
        datasources = json.loads(columns, object_pairs_hook=OrderedDict)
        tables = datasources.get(self._datasource, {})
        widths = tables.get(table, {})
        return widths

    def _getRecipientColumnsWidth(self):
        columns = self._configuration.getByName('MergerGrid2Columns')
        datasources = json.loads(columns, object_pairs_hook=OrderedDict)
        queries = datasources.get(self._datasource, {})
        widths = queries.get(self._query, {})
        return widths

    def _getColumnTitles(self, columns):
        titles = OrderedDict()
        for column in columns:
            titles[column] = column
        return titles

    def _getAddressColumns(self, name):
        table = self.Connection.getTables().getByName(name)
        columns = table.getColumns().getElementNames()
        return columns[:self._maxcolumns]

    def _getRecipientColumns(self):
        composer = self._getQueryComposer()
        columns = composer.getColumns().getElementNames()
        return columns[:self._maxcolumns]

    def _setAddressBook(self, table, initAddress):
        self._addressbook = table
        composer = self._getTableComposer(table)
        columns = self.getTableColumns(table)
        orders = composer.getOrderColumns().createEnumeration()
        initAddress(columns, orders)
        #filters = composer.getStructuredFilter()
        #command = self._getComposerCommand(table)
        #composer.setQuery(command)
        #composer.setStructuredFilter(filters)
        self._setRowSetCommand(self._address, composer)

    def _setAddressColumn(self, table, columns, reset):
        if reset:
            columns = self._getAddressColumns(table)
        titles = self._getColumnTitles(columns)
        self._column1.setColumnModel(self._address, titles, reset)

    def _setRecipientColumn(self, columns, reset):
        if reset:
            columns = self._getRecipientColumns()
        titles = self._getColumnTitles(columns)
        self._column2.setColumnModel(self._recipient, titles, reset)

    def _setAddressOrder(self, table, orders, ascending):
        composer = self._getTableComposer(table)
        self._setComposerOrder(composer, orders, ascending)
        self._setRowSetCommand(self._address, composer)

    def _setRecipientOrder(self, orders, ascending):
        composer = self._getQueryComposer()
        self._setComposerOrder(composer, orders, ascending)
        self._setRowSetCommand(self._recipient, composer)

    def _setComposerOrder(self, composer, orders, ascending):
        olds, news = self._getComposerOrder(composer, orders)
        composer.Order = ''
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
        composer = self._getQueryComposer()
        filters = self._getRowSetFilters(self._address, composer, rows, True)
        composer.setStructuredFilter(filters)
        self._setRowSetCommand(self._recipient, composer)

    def _removeItem(self, rows):
        composer = self._getQueryComposer()
        filters = self._getRowSetFilters(self._recipient, composer, rows, False)
        composer.setStructuredFilter(filters)
        self._setRowSetCommand(self._recipient, composer)

    def _getRowSetFilters(self, rowset, composer, rows, add):
        indexes = self._getIndexes(composer)
        filters = self._getComposerFilters(composer)
        for row in rows:
            rowset.absolute(row +1)
            filter = []
            for index in indexes:
                i = rowset.findColumn(index)
                value = getValueFromResult(rowset, i)
                filter.append((index, value))
            filter = tuple(filter)
            if add:
                self._addFilters(filters, filter)
            else:
                self._removeFilters(filters, filter)
        return self._getStructuredFilters(filters)

    def _getComposerFilters(self, composer):
        filters = []
        for f in composer.getStructuredFilter():
            filter = []
            for property in f:
                value = property.Value
                if isinstance(value, string_types):
                    value = value.strip("'")
                filter.append((property.Name, value))
            filters.append(tuple(filter))
        return filters

    def _addFilters(self, filters, filter):
        if filter not in filters:
            filters.append(filter)

    def _removeFilters(self, filters, filter):
        if filter in filters:
            filters.remove(filter)

    def _getStructuredFilters(self, filters):
        structured = []
        for filter in filters:
            properties = []
            for name, value in filter:
                operator = EQUAL if value != 'IS NULL' else SQLNULL
                property = getPropertyValue(name, value, None, operator)
                properties.append(property)
            properties = tuple(properties)
            structured.append(properties)
        structured = tuple(structured)
        return structured

    def _setDocumentRecord(self):
        try:
            dispatch = None
            frame = self._doc.getCurrentController().Frame
            flag = uno.getConstantByName('com.sun.star.frame.FrameSearchFlag.SELF')
            if self._doc.supportsService('com.sun.star.text.TextDocument'):
                url = getUrl(self._ctx, '.uno:DataSourceBrowser/InsertContent')
                dispatch = frame.queryDispatch(url, '_self', flag)
            elif self._doc.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
                url = getUrl(self._ctx, '.uno:DataSourceBrowser/InsertColumns')
                dispatch = frame.queryDispatch(url, '_self', flag)
            if dispatch is not None:
                descriptor = self._getDataDescriptor()
                dispatch.dispatch(url, descriptor)
        except Exception as e:
            print("MergerModel._setDocumentRecord() ERROR: %s" % traceback.print_exc())

    def _getDataDescriptor(self):
        properties = {'DataSourceName': self._recipient.DataSourceName,
                      'ActiveConnection': self._recipient.ActiveConnection,
                      'Command': self._recipient.Command,
                      'CommandType': self._recipient.CommandType,
                      'Cursor': self._recipient,
                      'Selection': (self._row, ),
                      'BookmarkSelection': False}
        descriptor = getPropertyValueSet(properties)
        return descriptor

# MergerModel StringRessoure methods
    def getPageStep(self, id):
        resource = self._getPageStep(id)
        step = self.resolveString(resource)
        return step

    def getPageTitle(self, id):
        resource = self._getPageTitle(id)
        title = self.resolveString(resource)
        return title

    def getTabTitle(self, id):
        resource = self._getTabTitleResource(id)
        return self.resolveString(resource)

    def getTabTip(self, id):
        resource = self._getTabTipResource(id)
        return self.resolveString(resource)

    def getEmailLabel(self):
        resource = self._getEmailResource()
        return self.resolveString(resource)

    def getIndexLabel(self):
        resource = self._getIndexResource()
        return self.resolveString(resource)

# MergerModel StringRessoure private methods
    def _getPageStep(self, id):
        return 'MergerPage%s.Step' % id

    def _getPageTitle(self, pageid):
        return 'MergerPage%s.Title' % pageid

    def _getTabTitleResource(self, id):
        return 'MergerTab%s.Title' % id

    def _getTabTipResource(self, id):
        return 'MergerTab%s.Tab.ToolTip' % id

    def _getEmailResource(self):
        return 'MergerPage1.Label11.Label.%s' % int(self._similar)

    def _getIndexResource(self):
        return 'MergerPage1.Label12.Label.%s' % int(self._similar)

# Procedures called internally
    # RowSet private methods
    def _getRowSet(self, command, fetchsize):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.CommandType = command
        rowset.FetchSize = fetchsize
        return rowset

    def _initRowSet(self):
        self._address.ActiveConnection = self.Connection
        self._recipient.ActiveConnection = self.Connection

    def _setRowSetCommand(self, rowset, composer):
        rowset.Command = composer.getQuery()
        rowset.execute()

    # DataSource private methods
    def _setDataSource(self, datasource, progress, setDataSource):
        progress(10)
        step = 2
        queries = label1 = label2 = msg = None
        sleep(0.2)
        progress(20)
        database = self._getDatabase(datasource)
        progress(30)
        try:
            if database.IsPasswordRequired:
                service = 'com.sun.star.task.InteractionHandler'
                handler = createService(self._ctx, service)
                connection = database.connectWithCompletion(handler)
            else:
                connection = database.getConnection('', '')
        except SQLException as e:
            msg = e.Message
        else:
            progress(40)
            self._datasource = datasource
            self._statement = connection.createStatement()
            progress(50)
            self._setTablesInfos()
            progress(60)
            self._initRowSet()
            progress(70)
            if datasource in self._composers:
                composers = self._composers[datasource]
            else:
                composers = {}
                self._composers[datasource] = composers
            progress(80)
            queries = self._getQueries(composers)
            progress(90)
            label1 = self.getEmailLabel()
            label2 = self.getIndexLabel()
            step = 3
        progress(100)
        setDataSource(step, queries, self._tables, label1, label2, msg)

    def _getDatabase(self, datasource):
        database = None
        if self._dbcontext.hasByName(datasource):
            database = self._dbcontext.getByName(datasource)
        return database

    def _setTablesInfos(self):
        tables = self.Connection.getTables()
        self._tables = tables.getElementNames()
        columns = tables.getByIndex(0).getColumns().getElementNames()
        for index in range(1, tables.getCount()):
            table = tables.getByIndex(index)
            if columns != table.getColumns().getElementNames():
                self._similar = False
                break
        else:
            self._similar = True

    # Composer private methods
    def _getComposers(self):
        return self._composers[self._datasource]

    def _getTableComposer(self, table):
        name = self._getComposerName(table)
        composers = self._getComposers()
        if name in composers:
            composer = composers[name]
        else:
            composer = self._createComposer()
            command = self._getComposerCommand(table)
            composer.setQuery(command)
            composers[name] = composer
        return composer

    def _getComposerNames(self):
        names = []
        for table in self._tables:
            name = self._getComposerName(table)
            names.append(name)
        return names

    def _getComposerName(self, table):
        name = '%s.%s' % (self._datasource, table)
        return name

    def _createComposer(self):
        service = 'com.sun.star.sdb.SingleSelectQueryComposer'
        composer = self.Connection.createInstance(service)
        return composer

    def _getComposerCommand(self, table):
        command = 'SELECT * FROM "%s"' % table
        return command

    def _getQueryComposer(self):
        composers = self._getComposers()
        composer = composers[self._query]
        return composer

    def _setComposerEmails(self, emails):
        filter = self._getEmailFilters(emails)
        tables = self._getComposerNames()
        for table in tables:
            composer = self._getTableComposer(table)
            composer.setStructuredFilter(filter)

    def _getComposerIndexes(self, composers):
        indexes = None
        if self._similar:
            queries = self._getQueries(composers)
            for query in queries:
                composer = composers[query]
                indexes = self._getIndexes(composer)
                break
        return indexes

    def _setComposerIndexes(self, indexes):
        filter = self._getIndexFilters(indexes)
        composers = self._getComposers()
        queries = self._getQueries(composers)
        for query in queries:
            composer = composers[query]
            composer.setStructuredFilter(filter)

    # Table private methods
    def _getTableColumns(self, table):
        table = self.Connection.getTables().getByName(table)
        columns = table.getColumns().getElementNames()
        return columns

    # Query private methods
    def _getQueries(self, composers):
        names = self._getComposerNames()
        queries = [name for name in composers if name not in names]
        return tuple(queries)

    # Email private methods
    def _getEmails(self, composer):
        emails = []
        if composer is not None:
            filters = composer.getStructuredFilter()
            for filter in filters:
                if len(filter) > 0:
                    emails.append(filter[0].Name)
        return emails

    def _setEmails(self, composer, emails):
        filter = self._getEmailFilters(emails)
        composer.setStructuredFilter(filter)

    def _getEmailFilters(self, emails):
        filters = []
        for email in emails:
            filter = getPropertyValue(email, 'IS NOT NULL', None, NOT_SQLNULL)
            filters.append((filter, ))
        return tuple(filters)

    # Index private methods
    def _getIndexes(self, composer):
        indexes = []
        filters = composer.getStructuredFilter()
        if len(filters) > 0:
            for filter in filters[0]:
                indexes.append(filter.Name)
        return indexes

    def _setIndexes(self, composer, indexes):
        filter = self._getIndexFilters(indexes)
        composer.setStructuredFilter(filter)

    def _getIndexFilters(self, indexes):
        filters = []
        for index in indexes:
            filter = getPropertyValue(index, 'IS NULL', None, SQLNULL)
            filters.append(filter)
        return (tuple(filters), )






    def _getQuery(self, queries, name, create):
        print("MergerModel._getQuery() '%s'" % (name, ))
        if queries.hasByName(name):
            query = queries.getByName(name)
        else:
            service = 'com.sun.star.sdb.QueryDefinition'
            query = createService(self._ctx, service)
            if create:
                queries.insertByName(name, query)
            else:
                table = self.Connection.getTables().getByIndex(0)
                column = table.getColumns().getByIndex(0).Name
                format = {'Table': table.Name, 'Column': column}
                query.Command = self._getQueryCommand(format)
                #query.UpdateTableName = table
        return query

    def _getDocumentName(self):
        url = None
        location = self._doc.getLocation()
        if location:
            url = getUrl(self._ctx, location)
        return None if url is None else url.Name
