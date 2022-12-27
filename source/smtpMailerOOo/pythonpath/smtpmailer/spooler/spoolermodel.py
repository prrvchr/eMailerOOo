#!
# -*- coding: utf-8 -*-

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

from com.sun.star.sdb.CommandType import TABLE

from ..grid import GridManager

from ..unotool import createService
from ..unotool import getConfiguration
from ..unotool import getPathSettings
from ..unotool import getStringResource

from ..dbtool import getValueFromResult

from ..logger import getMessage
from ..logger import logMessage

from ..configuration import g_identifier
from ..configuration import g_extension
from ..configuration import g_fetchsize

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
        self._rowset = self._getRowSet()
        self._grid = None
        self._configuration = getConfiguration(ctx, g_identifier, True)
        self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._resources = {'Title': 'SpoolerDialog.Title',
                           'State': 'SpoolerDialog.Label2.Label.%s',
                           'TabTitle': 'SpoolerTab%s.Title'}
        self._resource = 'SpoolerTab1.Grid1.Column.%s'

    @property
    def DataSource(self):
        return self._datasource

    @property
    def Path(self):
        return self._path
    @Path.setter
    def Path(self, path):
        self._path = path

# SpoolerModel getter methods
    def getGridSelectedRows(self):
        return self._grid.getSelectedRows()

    def isDisposed(self):
        return self._diposed

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
    def initSpooler(self, *args):
        Thread(target=self._initSpooler, args=args).start()

    def setGridData(self, rowset):
        self._grid.setRowSetData(rowset)

    def dispose(self):
        self._diposed = True
        if self._grid is not None:
            self._grid.dispose()
        self.DataSource.dispose()

    def saveGrid(self):
        self._grid.saveColumnSettings()

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

# SpoolerModel private getter methods
    def _getRowSet(self):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.CommandType = TABLE
        rowset.FetchSize = g_fetchsize
        return rowset

    def _getRowsJobs(self, rows):
        jobs = []
        i = self._rowset.findColumn('JobId')
        for row in rows:
            self._rowset.absolute(row +1)
            jobs.append(getValueFromResult(self._rowset, i))
        return tuple(jobs)

    def _getQueryTable(self):
        table = self._configuration.getByName('SpoolerTable')
        return table

# SpoolerModel private setter methods
    def _initSpooler(self, possize, parent, listener, initView):
        self.DataSource.waitForDataBase()
        self._rowset.ActiveConnection = self.DataSource.DataBase.Connection
        self._rowset.Command = self._getQueryTable()
        self._grid = GridManager(self._ctx, self._rowset, parent, possize, 'JobId', 'Spooler', self._resource, 13)
        self._grid.addSelectionListener(listener)
        initView(self._rowset)
        # TODO: GridColumn and GridModel needs a RowSet already executed!!!
        self.executeRowSet()
