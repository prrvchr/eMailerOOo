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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE


from smtpserver import GridModel

from unolib import getPropertyValueSet
from unolib import createService

from smtpserver.wizard import getOrders
from smtpserver.wizard import getOrder
from smtpserver.dbtools import getValueFromResult

from smtpserver import g_extension
from smtpserver import g_fetchsize

from smtpserver import logMessage
from smtpserver import getMessage

import time
import traceback


class DataBase(unohelper.Base):
    def __init__(self, ctx):
        print("DataBase.__init__() 1")
        self.ctx = ctx
        self._orders = []
        self._statement = None
        self._address = self._getRowSet()
        self._recipient = self._getRowSet()
        self._addressModel = GridModel(self._address)
        self._recipientModel = GridModel(self._recipient)
        print("DataBase.__init__() 2")

    @property
    def Connection(self):
        return self._statement.getConnection()
    @property
    def Address(self):
        return self._addressModel
    @property
    def Recipient(self):
        return self._recipientModel

    def getRecipient(self):
        return self._recipient

# Procedures called by the DataSource
    def setDataSource(self, dbcontext, progress, datasource, names, callback):
        time.sleep(0.2)
        step = 2
        tables, table, emails, keys, msg = None, None, None, None, None
        database = self._getDatabase(dbcontext, datasource)
        progress(20)
        try:
            if database.IsPasswordRequired:
                handler = createService(self.ctx, 'com.sun.star.task.InteractionHandler')
                connection = database.connectWithCompletion(handler)
            else:
                connection = database.getConnection('', '')
        except SQLException as e:
            query = None
            msg = e.Message
        else:
            progress(30)
            self._statement = connection.createStatement()
            document, form = self._getForm(False)
            progress(40)
            emails = self._getDocumentList(document, 'EmailColumns')
            keys = self._getDocumentList(document, 'IndexColumns')
            progress(50)
            if form is not None:
                form.close()
            progress(60)
            tbls = self.Connection.getTables()
            query = self._getQueryComposer(progress, names)
            tables = tbls.getElementNames()
            table = query.getTables().getByIndex(0)
            step = 3
        progress(100)
        callback(step, tables, table, emails, keys, msg)
        time.sleep(0.2)
        if query is not None:
            self._orders = getOrders(query.getOrder())
            filter = query.getFilter()
            self._initRowSet(datasource, tbls.getByIndex(0), filter)

    def initPage2(self, keys):
        print("DataBase.initPage2() 1: '%s' - %s" % (self._recipient.Filter, keys))
        self._setRowSetFilter(keys)
        self._executeRowSet()
        print("DataBase.initPage2() 2: '%s'" % self._recipient.Filter)

    def setRowSet(self, loaded, address, recipient, emails, keys):
        if not loaded:
            address = self.Connection.getTables().getByIndex(0).Name
        self._address.Filter = self._getNotNullFilter(address, emails)
        column = self._getRowSetColumns(keys)
        self._recipient.Command = self._getRowSetCommand(recipient, column)
        if loaded:
            self._setRowSetFilter(keys)
            self._executeRowSet()
        return loaded

    def getTableColumns(self, name):
        columns = self.Connection.getTables().getByName(name).getColumns().getElementNames()
        return columns

    def getTableName(self):
        return self.Connection.getTables().getByIndex(0).Name

    def getAddressRows(self):
        return range(self._address.RowCount)

    def getOrderIndex(self, columns):
        index = []
        for column in columns:
            if column in self._orders:
                index.append(columns.index(column))
        return tuple(index)

    def isConnected(self):
        return self._statement is not None

    def getDataSource(self):
        return self.Connection.getParent().DatabaseDocument.DataSource

    def executeRecipient(self, indexes, filters=(), filter=''):
        if len(filters) != 0:
            filter = ' OR '.join(filters)
        else:
            filter = self._getNullFilter(indexes)
        self._recipient.ApplyFilter = False
        self._recipient.Filter = filter
        self._recipient.ApplyFilter = True
        print("DataBase.executeRecipient()1 ************** %s" % self._recipient.ActiveCommand)
        self._recipient.execute()
        print("DataBase.executeRecipient()2 ************** %s" % filter)

    def executeAddress(self, emails, table, keys):
        column = self._getRowSetColumns(keys)
        self._address.Command = self._getRowSetCommand(table, column)
        self._address.Filter = self._getNotNullFilter(table, emails, True)
        self._address.ApplyFilter = True
        self._address.execute()

    # TODO: XRowset.Order should be treated as a stack where:
    # TODO: adding is done at the end and removing will keep order.
    def executeRowSet(self, address, recipient, columns, keys):
        print("DataBase.executeRowSet() 1")
        orders = getOrders(self._recipient.Order)
        print("DataBase.executeRowSet() 2: %s - %s" % (orders, columns))
        for order in reversed(self._orders):
            if order not in columns:
                self._orders.remove(order)
        for column in columns:
            if column not in self._orders:
                self._orders.append(column)
        print("DataBase.executeRowSet() 3: %s" % (self._orders, ))
        self._setRowSetCommand(address, recipient, keys)
        self._setRowSetOrder()
        self._executeRowSet()
        print("DataBase.executeRowSet() 4")

    def getAddressFilters(self, indexes, rows, recipients):
        filters = []
        for row in rows:
            self._address.absolute(row + 1)
            filter = self._getFilters(self._address, indexes)
            if filter not in recipients:
                filters.append(filter)
        print("DataBase.getAddressFilters() %s - %s)" % (self._address.RowCount, filters))
        return tuple(filters)

    def getRecipientFilters(self, indexes, rows=()):
        filters = []
        if self._recipient.RowCount > 0:
            self._recipient.beforeFirst()
            while self._recipient.next():
                row = self._recipient.Row -1
                if row not in rows:
                    filters.append(self._getFilters(self._recipient, indexes))
        print("DataBase.getRecipientFilters() %s - %s)" % (self._recipient.RowCount, filters))
        return tuple(filters)

# Procedures called internally
    def _getRowSet(self):
        rowset = createService(self.ctx, 'com.sun.star.sdb.RowSet')
        rowset.CommandType = COMMAND
        rowset.FetchSize = g_fetchsize
        return rowset

    def _getDatabase(self, dbcontext, datasource):
        database = None
        if dbcontext.hasByName(datasource):
            database = dbcontext.getByName(datasource)
        return database

    def _getForm(self, create, name='smtpMailerOOo'):
        doc, form = None, None
        forms = self.Connection.getParent().DatabaseDocument.getFormDocuments()
        if forms.hasByName(name):
            form = forms.getByName(name)
        elif create:
            form = self._createForm(forms, name)
        if form is not None:
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

    def _getDocumentList(self, document, property):
        items = ()
        value = self._getDocumentValue(document, property)
        if value is not None:
            items = tuple(value.split(','))
        return items

    def _getDocumentValue(self, document, property, default=None):
        value = default
        print("DataBase._getDocumentValue() %s" % (property, ))
        if document is not None:
            properties = document.DocumentProperties.UserDefinedProperties
            if properties.PropertySetInfo.hasPropertyByName(property):
                print("DataBase._getDocumentValue() getProperty")
                value = properties.getPropertyValue(property)
            elif default is not None:
                self._setDocumentValue(document, property, default)
        return value

    def _setDocumentValue(self, document, property, value):
        print("DataBase._setDocumentValue() %s - %s" % (property, value))
        properties = document.DocumentProperties.UserDefinedProperties
        if properties.PropertySetInfo.hasPropertyByName(property):
            print("DataBase._setDocumentValue() setProperty")
            properties.setPropertyValue(property, value)
        else:
            print("DataBase._setDocumentValue() addProperty")
            properties.addProperty(property,
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.MAYBEVOID') +
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.BOUND') +
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.REMOVABLE') +
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.MAYBEDEFAULT'),
            value)

    def _getQueryComposer(self, progress, names):
        composer = self.Connection.createInstance('com.sun.star.sdb.SingleSelectQueryComposer')
        progress(70)
        query = self._getQuery(names, False)
        progress(80)
        composer.setQuery(query.Command)
        print("DataBase._getQueryComposer() %s - %s" % (query.UpdateTableName, query.Command))
        progress(90)
        return composer

    def _getQuery(self, names, create, default=g_extension):
        print("DataBase._getQuery() '%s'" % (names, ))
        queries = self.getDataSource().getQueryDefinitions()
        for name in names:
            if queries.hasByName(name):
                query = queries.getByName(name)
                break
        else:
            query = createService(self.ctx, 'com.sun.star.sdb.QueryDefinition')
            if create:
                queries.insertByName(default, query)
            else:
                table = self.Connection.getTables().getByIndex(0)
                column = table.getColumns().getByIndex(0).Name
                format = {'Table': table.Name, 'Column': column}
                query.Command = self._getQueryCommand(format)
                query.UpdateTableName = table.Name
        return query

    def _initRowSet(self, datasource, table, filter):
        self._address.DataSourceName = datasource
        self._recipient.DataSourceName = datasource
        column = getOrder(self._orders)
        self._address.Command = self._getRowSetCommand(table.Name, column)
        self._recipient.Filter = filter
        self._setRowSetOrder()

    def _setRowSetCommand(self, address, recipient, keys):
        column = self._getRowSetColumns(keys)
        self._address.Command = self._getRowSetCommand(address, column)
        self._recipient.Command = self._getRowSetCommand(recipient, column)

    def _setRowSetOrder(self):
        order = getOrder(self._orders)
        self._address.Order = order
        self._recipient.Order = order

    def _setRowSetFilter(self, indexes):
        if self._recipient.Filter == '':
            self._recipient.ApplyFilter = False
            self._recipient.Filter = self._getNullFilter(indexes)
            self._recipient.ApplyFilter = True

    def _executeRowSet(self):
        self._address.execute()
        self._recipient.execute()

    def _addFilter(self, filters, any=False):
        separator = ' OR ' if any else ' AND '
        filter = separator.join(filters)
        if len(filters) > 1:
            filter = '(%s)' % filter
        return filter

    def _getFilters(self, rowset, indexes):
        filters = []
        columns = rowset.Columns.ElementNames
        for column in indexes:
            if column in columns:
                filter = '"%s"' % column
                i = rowset.findColumn(column)
                filter = "%s = '%s'" % (filter, getValueFromResult(rowset, i))
                filters.append(filter)
        return self._addFilter(filters)

    def _getNullFilter(self, indexes):
        filters = []
        for column in indexes:
            filter = '"%s" IS NULL' % column
            filters.append(filter)
        return self._addFilter(filters)

    def _getNotNullFilter(self, table, emails, any=False):
        filters = []
        columns = self.Connection.getTables().getByName(table).getColumns().getElementNames()
        for column in emails:
            if column in columns:
                filters.append('"%s" IS NOT NULL' % column)
        filter = self._addFilter(filters, any)
        print("PageModel._getFilter() %s" % filter)
        return filter

    def _getQueryCommand(self, format):
        query = 'SELECT "%(Column)s" FROM "%(Table)s" ORDER BY "%(Column)s"' % format
        return query

    def _getRowSetColumns(self, keys):
        columns = []
        for key in keys:
            if key not in self._orders:
                columns.append(key)
        return getOrder(self._orders + columns)

    def _getRowSetCommand(self, table, column):
        query = 'SELECT %s FROM "%s"' % (column, table)
        return query
