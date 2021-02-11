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

from com.sun.star.uno import XWeak
from com.sun.star.uno import XAdapter
from com.sun.star.sdbc import XRowSetListener
from com.sun.star.awt.grid import XMutableGridDataModel

from unolib import createService

from ..dbtools import getValueFromResult
from ..wizardtools import getOrders

import traceback


class GridDataModel(unohelper.Base,
                    XWeak,
                    XAdapter,
                    XRowSetListener,
                    XMutableGridDataModel):
    def __init__(self, ctx, rowset, ratio, titles, width):
        self._listeners = []
        self._datalisteners = []
        self._order = ''
        self._columns = ()
        self._ratio = ratio
        self._titles = titles
        # TODO: To display a Grid without a scroll bar, 1 must be subtracted
        self._width = width -1
        self.RowCount = 0
        self.ColumnCount = 0
        service = 'com.sun.star.awt.grid.DefaultGridColumnModel'
        self.ColumnModel = createService(ctx, service)
        self._resultset = None
        rowset.addRowSetListener(self)

    # XWeak
    def queryAdapter(self):
        return self

    # XAdapter
    def queryAdapted(self):
        return self
    def addReference(self, reference):
        pass
    def removeReference(self, reference):
        pass

    # XCloneable
    def createClone(self):
        return self

    # XGridDataModel
    def getCellData(self, column, row):
        self._resultset.absolute(row +1)
        return getValueFromResult(self._resultset, column +1)
    def getCellToolTip(self, column, row):
        return self.getCellData(column, row)
    def getRowHeading(self, row):
        return row
    def getRowData(self, row):
        data = []
        self._resultset.absolute(row +1)
        for index in range(self.ColumnCount):
            data.append(getValueFromResult(self._resultset, index +1))
        return tuple(data)

    # XMutableGridDataModel
    def addRow(self, heading, data):
        pass
    def addRows(self, headings, data):
        pass
    def insertRow(self, index, heading, data):
        pass
    def insertRows(self, index, headings, data):
        pass
    def removeRow(self, index):
        pass
    def removeAllRows(self):
        pass
    def updateCellData(self, column, row, value):
        pass
    def updateRowData(self, indexes, rows, values):
        pass
    def updateRowHeading(self, index, heading):
        pass
    def updateCellToolTip(self, column, row, value):
        pass
    def updateRowToolTip(self, row, value):
        pass
    def addGridDataListener(self, listener):
        self._datalisteners.append(listener)
    def removeGridDataListener(self, listener):
        if listener in self._datalisteners:
            self._datalisteners.remove(listener)

    # XComponent
    def dispose(self):
        event = uno.createUnoStruct('com.sun.star.lang.EventObject')
        event.Source = self
        for listener in self._listeners:
            listener.disposing(event)
    def addEventListener(self, listener):
        self._listeners.append(listener)
    def removeEventListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

    # XRowSetListener
    def disposing(self, event):
        pass
    def cursorMoved(self, event):
        pass
    def rowChanged(self, event):
        pass
    def rowSetChanged(self, event):
        rowset = event.Source
        self._resultset = rowset.createResultSet()
        self._setRowSetData(rowset)

    # Private methods
    def _setRowSetData(self, rowset):
        rowcount = self.RowCount
        self.RowCount = rowset.RowCount
        print("GridDataModel._setRowSetData() RowSet.RowCount=%s" % rowset.RowCount)
        self.ColumnCount = rowset.getMetaData().getColumnCount()
        columns = rowset.getColumns().getElementNames()
        if self._columns != columns:
            self._columns = columns
            self._setColumnModel()

        self._updateRowSetData(rowcount)

        #if rowcount != self.RowCount:
        #    self._updateRowSetData(rowcount)
        #if rowcount != 0 and self.RowCount != 0:
        #    self._changeRowSetData(0, self.RowCount)

    def _setColumnModel(self):
        for i in range(self.ColumnModel.getColumnCount() -1, -1, -1):
            self.ColumnModel.removeColumn(i)
        ratio = 0
        for name in self._columns:
            ratio += self._ratio[name]
            column = self.ColumnModel.createColumn()
            column.Identifier = name
            column.Title = self._titles[name]
            column.DataColumnIndex = self._columns.index(name)
            self.ColumnModel.addColumn(column)
        # TODO: ColumnWidth should be assigned after all columns have been added to the ColumnModel 
        self._setColumnWidth(ratio)

    def _setColumnWidth(self, ratio):
        i = 0
        width = 0
        last = self.ColumnModel.getColumnCount() -1
        for column in self.ColumnModel.getColumns():
            flex = self._ratio[column.Identifier]
            column.Flexibility = flex
            if i == last:
                column.ColumnWidth = self._width - width
            else:
                column.ColumnWidth = self._width * flex // ratio
            i += 1
            width += column.ColumnWidth

    def _updateRowSetData(self, rowcount):
        if self.RowCount < rowcount:
            self._removeRowSetData(self.RowCount, rowcount -1)
            if self.RowCount > 0:
                self._changeRowSetData(0, self.RowCount -1)
        elif self.RowCount > rowcount:
            self._insertRowSetData(rowcount, self.RowCount -1)
            if rowcount > 0:
                self._changeRowSetData(0, rowcount -1)
        elif rowcount > 0:
            self._changeRowSetData(0, rowcount -1)

    def _updateRowSetData1(self, rowcount):
        if self.RowCount < rowcount:
            self._removeRowSetData(self.RowCount, rowcount -1)
        else:
            self._insertRowSetData(rowcount, self.RowCount -1)

    def _removeRowSetData(self, first, last):
        event = self._getGridDataEvent(first, last)
        previous = None
        for listener in self._datalisteners:
            if previous != listener:
                listener.rowsRemoved(event)
                previous = listener

    def _insertRowSetData(self, first, last):
        event = self._getGridDataEvent(first, last)
        previous = None
        for listener in self._datalisteners:
            if previous != listener:
                listener.rowsInserted(event)
                previous = listener

    def _changeRowSetData(self, first, last):
        event = self._getGridDataEvent(first, last)
        previous = None
        for listener in self._datalisteners:
            if previous != listener:
                listener.dataChanged(event)
                previous = listener

    def _getGridDataEvent(self, first, last):
        event = uno.createUnoStruct('com.sun.star.awt.grid.GridDataEvent')
        event.Source = self
        event.FirstColumn = 0
        event.LastColumn = self.ColumnCount -1
        event.FirstRow = first
        if first != -1:
           event.LastRow = last
        return event
