#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-25 https://prrvchr.github.io                                  ║
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

from .grid import GridManager as GridManagerBase

from .gridmodel import GridModel

from collections import OrderedDict
import json
import traceback


class GridManager(GridManagerBase):
    def __init__(self, ctx, url, window, quote, setting, selection, resource=None, maxi=None, multi=False):
        GridManagerBase.__init__(self, ctx, url, GridModel(), window, quote, setting, selection, resource, maxi, multi)

# GridManager setter methods
    def setDataModel(self, rowset, identifiers):
        datasource = self._getRowSetName(rowset)
        table = rowset.UpdateTableName
        changed = self._isDataSourceChanged(datasource, table)
        if changed:
            if self._isGridLoaded():
                self._saveWidths()
                self._saveOrders()
            # We can hide GridColumnHeader and reset GridDataModel
            # but after saving GridColumnModel Widths
            self._view.showGridColumnHeader(False)
            self._headers, self._indexes, self._types = self._getHeadersInfo(rowset.getMetaData(), identifiers)
            self._view.initColumns(self._url, self._headers, self._initColumnModel(datasource, table))
            self._table = table
            self._datasource = datasource
            self._view.showGridColumnHeader(True)
        sort = self.Model.getCurrentSortOrder()
        # XXX: Before changing the data, the sort order must be deactivated
        self.Model.removeColumnSort()
        notified = self._model.hasGridDataListener()
        if not notified:
            self.setGridVisible(False)
        self._model.resetRowSetData()
        self._model.setRowSetData(rowset)
        if not notified:
            self.setGridVisible(True)
        # XXX: After changing the data, the sort order must be activated if needed...
        if changed:
            index, ascending = self._getSavedOrders(datasource, table)
            if index != -1: 
                self.Model.sortByColumn(index, ascending)
            else:
                self.Model.removeColumnSort()
        elif sort.First != -1:
            self.Model.sortByColumn(sort.First, sort.Second)

    def resetDataModel(self):
        if self._isGridLoaded():
            self._saveWidths()
            self._saveOrders()
        self._model.resetRowSetData()
        self._table = None
        self._datasource = None

    def getGridPredicates(self):
        predicates = []
        for row in (range(self._model.RowCount)):
            predicates.append(json.dumps(self.getRowPredicates(row)))
        return tuple(predicates)

# GridManager private methods
    def _isDataSourceChanged(self, datasource, table):
        return self._datasource != datasource or self._table != table

    def _isGridLoaded(self):
        return self._datasource is not None

    def _getHeadersInfo(self, metadata, identifiers):
        headers = OrderedDict()
        indexes = OrderedDict([(identifier, -1) for identifier in identifiers])
        types = {}
        for i in range(metadata.getColumnCount()):
            name = metadata.getColumnLabel(i +1)
            title = self._getColumnTitle(name)
            if name in identifiers:
                indexes[name] = i
                types[name] = metadata.getColumnType(i + 1)
            headers[name] = title
        return headers, indexes, types

    def _getRowSetName(self, rowset):
        datasource = rowset.DataSourceName
        return datasource if datasource else rowset.ActiveConnection.Parent.Name
