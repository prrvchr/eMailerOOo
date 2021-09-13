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

from smtpmailer import GridModel
from smtpmailer import ColumnModel

from smtpmailer import createService
from smtpmailer import getConfiguration
from smtpmailer import getMessage
from smtpmailer import getPathSettings
from smtpmailer import getStringResource
from smtpmailer import getValueFromResult
from smtpmailer import logMessage
from smtpmailer import g_identifier
from smtpmailer import g_extension
from smtpmailer import g_fetchsize

from collections import OrderedDict
from threading import Thread
import validators
import json
import traceback


class SpoolerModel(unohelper.Base):
    def __init__(self, ctx, datasource):
        self._ctx = ctx
        self._diposed = False
        self._path = getPathSettings(ctx).Work
        self._datasource = datasource
        self._column = ColumnModel(ctx)
        self._rowset = self._getRowSet()
        self._composer = None
        self._configuration = getConfiguration(ctx, g_identifier, True)
        self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._resources = {'Title': 'SpoolerDialog.Title',
                           'State': 'SpoolerDialog.Label2.Label.%s',
                           'TabTitle': 'SpoolerTab%s.Title',
                           'GridColumn': 'SpoolerTab1.Grid1.Column.%s'}

    @property
    def DataSource(self):
        return self._datasource

    @property
    def Path(self):
        return self._path
    @Path.setter
    def Path(self, path):
        self._path = path

# SpoolerModel threaded callback methods
    def initSpooler(self, *args):
        Thread(target=self._initSpooler, args=args).start()

    def _initSpooler(self, initView):
        self.DataSource.waitForDataBase()
        command = self._getQueryCommand()
        self._composer = self._getQueryComposer(command)
        self._rowset.ActiveConnection = self.DataSource.DataBase.Connection
        self._rowset.Command = command
        titles = self._getQueryColumnTitles()
        orders = self._getQueryOrder()
        initView(titles, orders)

# SpoolerModel getter methods
    def isDisposed(self):
        return self._diposed

    def getGridModels(self, titles, width):
        data = GridModel(self._rowset)
        widths = self._getColumnsWidth()
        column = self._column.getModel(self._rowset, widths, titles, width, 2)
        return data, column

    def getDialogTitle(self):
        resource = self._resources.get('Title')
        return self._resolver.resolveString(resource)

    def getTabPageTitle(self, tabid):
        resource = self._resources.get('TabTitle') % tabid
        return self._resolver.resolveString(resource)

    def getSpoolerState(self, state):
        resource = self._resources.get('State') % state
        print("SpoolerModel.getSpoolerState() %s" % resource)
        return self._resolver.resolveString(resource)

# SpoolerModel setter methods
    def dispose(self):
        self._diposed = True
        self.DataSource.dispose()

    def save(self):
        widths = self._column.getWidths()
        self._saveGridColumn(widths)
        command = self._composer.getQuery()
        self._saveQueryCommand(command)
        self._configuration.commitChanges()

    def setGridColumnModel(self, titles, reset):
        self._column.setModel(self._rowset, titles, reset)

    def removeRows(self, rows):
        jobs = self._getRowsJobs(rows)
        print("SpoolerModel.removeRows() 1 %s" % (jobs,))
        if self.DataSource.deleteJob(jobs):
            print("SpoolerModel.removeRows() 2")
            self._rowset.execute()

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

# SpoolerModel private getter methods
    def _getRowSet(self):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.CommandType = COMMAND
        rowset.FetchSize = g_fetchsize
        return rowset

    def _getRowsJobs(self, rows):
        jobs = []
        i = self._rowset.findColumn('JobId')
        for row in rows:
            self._rowset.absolute(row +1)
            jobs.append(getValueFromResult(self._rowset, i))
        return tuple(jobs)

    def _getQueryComposer(self, command):
        service = 'com.sun.star.sdb.SingleSelectQueryComposer'
        composer = self.DataSource.DataBase.Connection.createInstance(service)
        # TODO: For the SingleSelectQueryComposer to be able to parse the Order, the ORDER BY
        # TODO: sort criteria must be in the SQL query.Command and not in the query.Order
        composer.setQuery(command)
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
            resource = self._resources.get('GridColumn') % column
            titles[column] = self._resolver.resolveString(resource)
        return titles

    def _getQueryCommand(self):
        command = self._configuration.getByName('SpoolerQueryCommand')
        return command

    def _getColumnsWidth(self):
        columns = self._configuration.getByName('SpoolerGridColumns')
        widths = json.loads(columns, object_pairs_hook=OrderedDict)
        return widths

# SpoolerModel private setter methods
    def _saveGridColumn(self, widths):
        columns = json.dumps(widths)
        self._configuration.replaceByName('SpoolerGridColumns', columns)

    def _saveQueryCommand(self, command):
        self._configuration.replaceByName('SpoolerQueryCommand', command)
