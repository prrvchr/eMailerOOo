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

from unolib import createService
from unolib import getPathSettings
from unolib import getStringResource

from .griddatamodel import GridDataModel

from smtpserver import g_identifier
from smtpserver import g_extension
from smtpserver import g_fetchsize

from smtpserver import logMessage
from smtpserver import getMessage

import validators
import traceback


class SpoolerModel(unohelper.Base):
    def __init__(self, ctx, datasource):
        self._ctx = ctx
        self._path = getPathSettings(ctx).Work
        self._datasource = datasource
        self._ratio = {'Id': 4, 'State': 5, 'Subject': 15, 'Sender': 20,
                       'Recipient': 20, 'Document': 25, 'TimeStamp': 11}
        self._rowset = self._getRowSet()
        self._stringResource = getStringResource(ctx, g_identifier, g_extension)

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

    def getGridDataModel(self, width):
        titles = self._getGridTitles()
        return GridDataModel(self._ctx, self._rowset, self._ratio, titles, width)

    def initRowSet(self, callback):
        self._datasource.initRowSet(callback)

    def setRowSet(self):
        database = self._datasource.DataBase
        #mri = createService(self._ctx, 'mytools.Mri')
        #mri.inspect(database.Connection)
        #self._rowset.DataSourceName = database.getDataSource().Name
        self._rowset.ActiveConnection = database.Connection
        #self._rowset.Command = query.Command
        self._rowset.Command = database.getRowSetCommand()
        #self._rowset.Order = query.Order
        # TODO: if RowSet.Filter is not assigned then RowSet.RowCount is 1 anyway
        #self._rowset.Filter = '1=0'
        print("SpoolerModel.setRowSet() RowSet.Command: %s" % self._rowset.Command)
        self._rowset.execute()

    def getTableColumns(self, name):
        columns = self._datasource.DataBase.Connection.getTables().getByName(name).getColumns().getElementNames()
        return columns

    def getRowSet(self):
        return self._rowset

    def executeRowSet(self):
        self._rowset.ApplyFilter = True
        self._rowset.ApplyFilter = False
        self._rowset.execute()

    def getQueryComposer(self, query):
        database = self.DataSource.DataBase
        service = 'com.sun.star.sdb.SingleSelectQueryComposer'
        composer = database.Connection.createInstance(service)
        composer.setQuery(query.Command)
        print("SpoolerModel._getQueryComposer() %s" % (query.Command))
        return composer

    def getQuery(self, name='Spooler'):
        database = self.DataSource.DataBase
        print("SpoolerModel._getQuery() '%s'" % name)
        queries = database.getDataSource().getQueryDefinitions()
        if queries.hasByName(name):
            query = queries.getByName(name)
            print("SpoolerModel._getQuery() %s exist" % name)
        else:
            service = 'com.sun.star.sdb.QueryDefinition'
            query = createService(self._ctx, service)
            query.Command = database.getQueryCommand()
            queries.insertByName(name, query)
            print("SpoolerModel._getQuery() %s don't exist" % name)
        return query

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

    def _getTabPageResource(self, id):
        return 'SpoolerDialog.Tab1.Page%s.Title' % id

    def _getGridColumnResource(self, column):
        return 'SpoolerDialog.Grid1.Column.%s' % column

    def _getSpoolerStateResource(self, state):
        return 'SpoolerDialog.Label2.Label.%s' % state

# SpoolerModel private methods
    def _getRowSet(self):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.CommandType = COMMAND
        rowset.FetchSize = g_fetchsize
        return rowset

    def _getGridTitles(self):
        titles = {}
        for column in self._ratio:
            resource = self._getGridColumnResource(column)
            titles[column] = self.resolveString(resource)
        return titles
