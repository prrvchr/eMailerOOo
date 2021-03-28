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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE
from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY

from unolib import createService
from unolib import getPathSettings
from unolib import getStringResource
from unolib import getConfiguration

from smtpserver.grid import GridModel
from smtpserver.grid import ColumnModel

from smtpserver import g_identifier
from smtpserver import g_extension
from smtpserver import g_fetchsize

from smtpserver import logMessage
from smtpserver import getMessage

from collections import OrderedDict
import validators
import json
from threading import Thread
import traceback


class SpoolerModel(unohelper.Base):
    def __init__(self, ctx, datasource):
        self._ctx = ctx
        self._path = getPathSettings(ctx).Work
        self._datasource = datasource
        self._column = ColumnModel(ctx)
        self._rowset = self._getRowSet()
        self._stringResource = getStringResource(ctx, g_identifier, g_extension)
        self._configuration = getConfiguration(ctx, g_identifier, True)
        self._composer = None

    @property
    def DataSource(self):
        return self._datasource

    @property
    def Path(self):
        return self._path
    @Path.setter
    def Path(self, path):
        self._path = path

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

    def initSpooler(self, *args):
        Thread(target=self._initSpooler, args=args).start()

    def _initSpooler(self, initView):
        self.DataSource.waitForDataBase()
        query = self._getQuery()
        self._composer = self._getQueryComposer(query)
        self._rowset.ActiveConnection = self.DataSource.DataBase.Connection
        self._rowset.Command = query.Command
        titles = self._getQueryColumnTitles()
        orders = self._getQueryOrder()
        initView(titles, orders)

    def getGridModels(self, titles, width):
        data = GridModel(self._rowset)
        widths = self._getColumnsWidth()
        column = self._column.getColumnModel(self._rowset, widths, titles, width, 2)
        return data, column

    def setGridColumnModel(self, titles, reset):
        self._column.setColumnModel(self._rowset, titles, reset)

    def saveGridColumn(self):
        widths = self._column.getColumnWidth()
        columns = json.dumps(widths)
        self._configuration.replaceByName('SpoolerGridColumns', columns)
        self._configuration.commitChanges()

    def saveQuery(self):
        query = self._getQuery()
        query.Command = self._composer.getQuery()
        self.DataSource.DataBase.storeDataBase()

    def executeRowSet(self):
        # TODO: If RowSet.Filter is not assigned then unassigned, RowSet.RowCount is always 1
        self._rowset.ApplyFilter = True
        self._rowset.ApplyFilter = False
        self._rowset.execute()

    def setRowSetOrder(self, orders, ascending):
        olds, news = self._getComposerOrder(orders)
        self._composer.Order = ''
        for order in olds:
            self._composer.appendOrderByColumn(order, order.IsAscending)
        columns = self._composer.getColumns()
        for order in news:
            column = columns.getByName(order)
            self._composer.appendOrderByColumn(column, ascending)
        self._rowset.Command = self._composer.getQuery()
        self._rowset.execute()

# SpoolerModel StringRessoure methods
    def getDialogTitle(self):
        resource = self._getDialogTitleResource()
        return self.resolveString(resource)

    def getTabPageTitle(self, id):
        resource = self._getTabPageResource(id)
        return self.resolveString(resource)

    def getSpoolerState(self, state):
        resource = self._getSpoolerStateResource(state)
        return self.resolveString(resource)

# SpoolerModel StringRessoure private methods
    def _getDialogTitleResource(self):
        return 'SpoolerDialog.Title'

    def _getSpoolerStateResource(self, state):
        return 'SpoolerDialog.Label2.Label.%s' % state

    def _getTabPageResource(self, id):
        return 'SpoolerTab%s.Title' % id

    def _getGridColumnResource(self, column):
        return 'SpoolerTab1.Grid1.Column.%s' % column

# SpoolerModel private methods
    def _getRowSet(self):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.CommandType = COMMAND
        rowset.FetchSize = g_fetchsize
        return rowset

    def _getQuery(self, name='SpoolerView'):
        queries = self.DataSource.DataBase.getDataSource().getQueryDefinitions()
        # TODO: For the SingleSelectQueryComposer to be able to parse the Order,
        # TODO: we need a sub query and not directly a DataBase View or Table query!!!
        self._setQuery(queries)
        if queries.hasByName(name):
            query = queries.getByName(name)
        else:
            command = self.DataSource.DataBase.getSpoolerViewQuery()
            query = self._createQuery(name, command)
            queries.insertByName(name, query)
        return query

    def _setQuery(self, queries, name='View'):
        if not queries.hasByName(name):
            command = self.DataSource.DataBase.getViewQuery()
            query = self._createQuery(name, command)
            queries.insertByName(name, query)

    def _createQuery(self, name, command):
        service = 'com.sun.star.sdb.QueryDefinition'
        query = createService(self._ctx, service)
        query.Command = command
        return query

    def _getQueryComposer(self, query):
        service = 'com.sun.star.sdb.SingleSelectQueryComposer'
        composer = self.DataSource.DataBase.Connection.createInstance(service)
        # TODO: For the SingleSelectQueryComposer to be able to parse the Order, the ORDER BY
        # TODO: sort criteria must be in the SQL query.Command and not in the query.Order
        composer.setQuery(query.Command)
        return composer

    def _getComposerOrder(self, news):
        olds = []
        enumeration = self._getQueryOrder()
        while enumeration.hasMoreElements():
            column = enumeration.nextElement()
            if column.Name in news:
                olds.append(column)
                news.remove(column.Name)
        return olds, news

    def _getQueryOrder(self):
        return self._composer.getOrderColumns().createEnumeration()

    def _getQueryColumnTitles(self):
        columns = self._composer.getColumns().getElementNames()
        titles = self._getColumnTitles(columns)
        return titles

    def _getColumnTitles(self, columns):
        titles = OrderedDict()
        for column in columns:
            resource = self._getGridColumnResource(column)
            titles[column] = self.resolveString(resource)
        return titles

    def _getColumnsWidth(self):
        columns = self._configuration.getByName('SpoolerGridColumns')
        widths = json.loads(columns, object_pairs_hook=OrderedDict)
        return widths
