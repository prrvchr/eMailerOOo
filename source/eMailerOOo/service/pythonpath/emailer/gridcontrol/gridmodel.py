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

from ..grid import GridModel as GridModelBase

from ..dbtool import getResultValue

import traceback


class GridModel(GridModelBase):
    def __init__(self, rowset=None):
        GridModelBase.__init__(self)
        self._rowset = rowset
        self._resultset, self._row, self._column = self._getRowsetData(rowset)

    @property
    def RowCount(self):
        return self._row
    @property
    def ColumnCount(self):
        return self._column

# XCloneable
    def createClone(self):
         return GridModel(self._rowset)

# XGridDataModel
    def getCellData(self, column, row):
        self._resultset.absolute(row + 1)
        return  getResultValue(self._resultset, column + 1)

    def getCellToolTip(self, column, row):
        return self.getCellData(column, row)

    def getRowData(self, row):
        data = []
        self._resultset.absolute(row +1)
        for column in range(self._column):
            data.append(getResultValue(self._resultset, column + 1))
        return tuple(data)

# GridModel setter methods
    def resetRowSetData(self):
        hasrow = self._row > 0
        self._row = 0
        if hasrow:
            self.removeRow(-1, -1)

    def setRowSetData(self, rowset):
        row = self._row
        self._resultset, self._row, self._column = self._getRowsetData(rowset)
        return row, self._row

    def removeRow(self, first, last):
        event = self._getGridDataEvent(first, last)
        for listener in self._listeners:
            listener.rowsRemoved(event)

    def insertRow(self, first, last):
        event = self._getGridDataEvent(first, last)
        for listener in self._listeners:
            listener.rowsInserted(event)

    def changeData(self, first, last):
        event = self._getGridDataEvent(first, last)
        for listener in self._listeners:
            listener.dataChanged(event)

# GridModel private methods
    def _getRowsetData(self, rowset):
        if rowset is None:
            return None, 0 , 0
        return rowset.createResultSet(), rowset.RowCount, rowset.getMetaData().getColumnCount()

    def _getGridDataEvent(self, first, last):
        event = uno.createUnoStruct('com.sun.star.awt.grid.GridDataEvent')
        event.Source = self
        event.FirstColumn = -1
        event.LastColumn = -1
        event.FirstRow = first
        event.LastRow = last
        return event

