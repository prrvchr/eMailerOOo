#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.uno import XWeak
from com.sun.star.uno import XAdapter
from com.sun.star.sdbc import XRowSetListener
from com.sun.star.awt.grid import XMutableGridDataModel

from unolib import createService

from .dbtools import getValueFromResult
from .wizardtools import getRowSetOrders

import traceback


class GridDataModel(unohelper.Base,
                    XWeak,
                    XAdapter,
                    XRowSetListener,
                    XMutableGridDataModel):
    def __init__(self, ctx, rowset):
        self._listeners = []
        self._datalisteners = []
        self._order = ''
        self.RowCount = self.ColumnCount = 0
        self.ColumnModel = createService(ctx, 'com.sun.star.awt.grid.DefaultGridColumnModel')
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
        self._resultset.absolute(row + 1)
        return getValueFromResult(self._resultset, column + 1)
    def getCellToolTip(self, column, row):
        return self.getCellData(column, row)
    def getRowHeading(self, row):
        return row
    def getRowData(self, row):
        data = []
        self._resultset.absolute(row + 1)
        for index in range(self.ColumnCount):
            data.append(getValueFromResult(self._resultset, index + 1))
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

    def _setRowSetData(self, rowset):
        print("GridDataModel._setRowSetData()1")
        rowcount = self.RowCount
        self.RowCount = rowset.RowCount
        metadata = rowset.getMetaData()
        self.ColumnCount = metadata.getColumnCount()
        if rowset.Order != self._order:
            self._setColumnModel(rowset, metadata)
        print("GridDataModel._setRowSetData()2")
        if rowcount != self.RowCount:
            self._updateRowSetData(rowcount)
        if rowcount != 0 and self.RowCount != 0:
            self._changeRowSetData(0, self.RowCount)
        print("GridDataModel._setRowSetData()3")

    def _setColumnModel(self, rowset, metadata):
        orders = getRowSetOrders(rowset)
        print("GridDataModel._setColumnModel()1")
        for i in range(self.ColumnModel.getColumnCount(), 0 ,-1):
            name = self.ColumnModel.getColumn(i - 1).Title
            if name in orders:
                orders.remove(name)
            else:
                self.ColumnModel.removeColumn(i - 1)
        truncated = False
        columns = rowset.getColumns()
        print("GridDataModel._setColumnModel()2")
        for name in orders:
            if not columns.hasByName(name):
                truncated = True
                continue
            index = rowset.findColumn(name)
            column = self.ColumnModel.createColumn()
            column.Title = name
            size = metadata.getColumnDisplaySize(index)
            column.MinWidth = size // 2
            column.DataColumnIndex = index - 1
            self.ColumnModel.addColumn(column)
        print("GridDataModel._setColumnModel()3")
        if truncated:
            orders = [column.Title for column in self.ColumnModel.getColumns()]
            order = '"%s"' % '", "'.join(orders) if len(orders) else ''
            self._order = rowset.Order = order
        else:
            self._order = rowset.Order
        print("GridDataModel._setColumnModel()4")

    def _updateRowSetData(self, rowcount):
        if self.RowCount < rowcount:
            self._removeRowSetData(self.RowCount, rowcount -1)
        else:
            self._insertRowSetData(rowcount, self.RowCount -1)

    def _removeRowSetData(self, first, last):
        try:
            print("GridDataModel._removeRowSetData()1")
            event = self._getGridDataEvent(first, last)
            print("GridDataModel._removeRowSetData()2")
            for listener in self._datalisteners:
                print("GridDataModel._removeRowSetData()3")
                listener.rowsRemoved(event)
            print("GridDataModel._removeRowSetData()4")
        except Exception as e:
            print("GridDataModel._removeRowSetData() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _insertRowSetData(self, first, last):
        event = self._getGridDataEvent(first, last)
        for listener in self._datalisteners:
            listener.rowsInserted(event)

    def _changeRowSetData(self, first, last):
        event = self._getGridDataEvent(first, last)
        for listener in self._datalisteners:
            listener.dataChanged(event)

    def _getGridDataEvent(self, first, last):
        event = uno.createUnoStruct('com.sun.star.awt.grid.GridDataEvent')
        event.Source = self
        event.FirstColumn = 0
        event.LastColumn = self.ColumnCount - 1
        event.FirstRow = first
        if first != -1:
           event.LastRow = last
        return event
