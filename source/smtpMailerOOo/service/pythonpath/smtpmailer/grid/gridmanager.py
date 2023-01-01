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

from .gridmanagerbase import GridManagerBase

from collections import OrderedDict
import traceback


class GridManager(GridManagerBase):
    def __init__(self, ctx, url, model, window, setting, selection, resource=None, maxi=None, multi=False):
        GridManagerBase.__init__(self, ctx, url, model, window, setting, selection, resource, maxi, multi)

# GridManager setter methods
    def setDataModel(self, rowset, identifiers):
        datasource = rowset.ActiveConnection.Parent.Name
        query = rowset.UpdateTableName
        changed = self._isDataSourceChanged(datasource, query)
        if changed:
            if self._isGridLoaded():
                self._saveWidths()
                self._saveOrders()
            # We can hide GridColumnHeader and reset GridDataModel
            # but after saving GridColumnModel Widths
            self._view.showGridColumnHeader(False)
            #self._model.resetRowSetData()
            self._headers, self._indexes, self._types = self._getHeadersInfo(rowset.getMetaData(), identifiers)
            self._view.initColumns(self._url, self._headers, self._initColumnModel(datasource, query))
            self._query = query
            self._datasource = datasource
            self._view.showGridColumnHeader(True)
        self._view.setGridVisible(False)
        self._model.setRowSetData(rowset)
        self._view.setGridVisible(True)
        if changed:
            self._model.sortByColumn(*self._getSavedOrders(datasource, query))

# GridManager private methods
    def _isDataSourceChanged(self, name, query):
        return self._datasource != name or self._query != query

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
                types[name] = metadata.getColumnType(i +1)
            headers[name] = title
        return headers, indexes, types

