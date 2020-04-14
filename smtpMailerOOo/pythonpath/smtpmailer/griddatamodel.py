#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.uno import XWeak
from com.sun.star.uno import XAdapter
from com.sun.star.sdbc import XRowSetListener
from com.sun.star.awt.grid import XMutableGridDataModel

from unolib import createService

from .configuration import g_default_columns

from .dbtools import getValueFromResult

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
        for column in range(self.ColumnCount):
            data.append(getValueFromResult(self._resultset, column + 1))
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
        if rowset.Order != self._order:
            self._setColumnModel(rowset)
        if self.RowCount > 0:
            self._removeRowSetData(rowset)
        self._insertRowSetData(rowset)

    def _setColumnModel(self, rowset):
        orders = rowset.Order.strip('"').split('","')
        for i in range(self.ColumnModel.getColumnCount(), 0 ,-1):
            name = self.ColumnModel.getColumn(i - 1).Title
            if name in orders:
                orders.remove(name)
            else:
                self.ColumnModel.removeColumn(i - 1)
        truncated = False
        columns = rowset.getColumns()
        metadata = rowset.getMetaData()
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
        if truncated:
            orders = [column.Title for column in self.ColumnModel.getColumns()]
            order = '"%s"' % '","'.join(orders) if len(orders) else ''
            self._order = rowset.Order = order
        else:
            self._order = rowset.Order

    def _removeRowSetData(self, rowset):
        self.RowCount = self.ColumnCount = 0
        event = self._getGridDataEvent()
        for listener in self._datalisteners:
            listener.rowsRemoved(event)

    def _insertRowSetData(self, rowset):
        self.RowCount = rowset.RowCount
        self.ColumnCount = rowset.getMetaData().getColumnCount()
        if self.RowCount > 0:
            event = self._getGridDataEvent(0)
            for listener in self._datalisteners:
                listener.rowsInserted(event)

    def _getGridDataEvent(self, first=-1):
        event = uno.createUnoStruct('com.sun.star.awt.grid.GridDataEvent')
        event.Source = self
        event.FirstColumn = event.FirstRow = first
        event.LastColumn = self.ColumnCount - 1
        event.LastRow = self.RowCount - 1
        return event
