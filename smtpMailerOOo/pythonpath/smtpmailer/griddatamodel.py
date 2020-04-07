#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.uno import XWeak
from com.sun.star.uno import XAdapter
from com.sun.star.sdbc import XRowSetListener
from com.sun.star.awt.grid import XMutableGridDataModel

from .dbtools import getValueFromResult

import traceback


class GridDataModel(unohelper.Base,
                    XWeak,
                    XAdapter,
                    XRowSetListener,
                    XMutableGridDataModel):
    def __init__(self, rowset):
        self._resultset = rowset.createResultSet()
        self._listeners = []
        self._datalisteners = []
        self.RowCount = self.ColumnCount = 0
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
        event = uno.createUnoStruct('com.sun.star.awt.grid.GridDataEvent')
        event.Source = self
        if self.RowCount > 0:
            self.RowCount = self.ColumnCount = 0
            event = self._setDataEvent(event)
            for listener in self._datalisteners:
                listener.rowsRemoved(event)
        self.RowCount = rowset.RowCount
        self.ColumnCount = self._resultset.getMetaData().getColumnCount()
        if self.RowCount > 0:
            event = self._setDataEvent(event, 0)
            for listener in self._datalisteners:
                listener.rowsInserted(event)

    def _setDataEvent(self, event, first=-1):
        event.FirstColumn = event.FirstRow = first
        event.LastColumn = self.ColumnCount - 1
        event.LastRow = self.RowCount - 1
        return event
