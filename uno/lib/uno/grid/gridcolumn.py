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

from .gridhandler import GridHandler

from ..unotool import createService
from ..unotool import getConfiguration
from ..unotool import getStringResource

from ..configuration import g_extension
from ..configuration import g_identifier

import json
from collections import OrderedDict
import traceback


class GridColumn(unohelper.Base):
    def __init__(self, ctx, rowset, resource=None, name=None, maximum=0, multi=False):
        self._ctx = ctx
        self._factor = 5
        self._rowset = rowset
        self._datasource = None
        self._resource = resource
        if resource is not None:
            self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._name = name
        self._widths = {}
        if name is not None:
            self._config = getConfiguration(ctx, g_identifier, True)
            self._widths = self._getConfigWidths()
        self._max = maximum
        self._multi = multi
        service = 'com.sun.star.awt.grid.DefaultGridColumnModel'
        self._column = createService(ctx, service)
        handler = GridHandler(self)
        rowset.addRowSetListener(handler)

# GridColumn getter methods
    def getColumnModel(self):
        return self._column

    def getColumnTitles(self):
        if self._resource is None:
            titles = self._getDefaultTitles()
        else:
            titles = self._getResourceTitles()
        return titles

    def getColumns(self):
        columns = []
        metadata = self._rowset.getMetaData()
        for index in range(metadata.getColumnCount()):
            column = metadata.getColumnLabel(index +1)
            columns.append(column)
        return columns

# GridColumn setter methods
    def setRowSetData(self, rowset):
            datasource = rowset.ActiveConnection.Parent.Name
            if datasource != self._datasource:
                self._initModel(datasource)
                self._datasource = datasource

    def saveColumnwidths(self):
        widths = None
        if self._name is not None:
            width = self._getColumnWidths()
            if not self._multi:
                widths = width
            elif self._datasource is not None:
                self._widths[self._datasource] = width
                widths = self._widths
            if widths is not None:
                columns = json.dumps(widths)
                self._config.replaceByName(self._name, columns)
                self._config.commitChanges()

    def setColumnModel(self, titles, reset):
        modified = False
        for index in range(self._column.getColumnCount() -1, -1, -1):
            column = self._column.getColumn(index)
            if column.Identifier not in titles:
                self._column.removeColumn(index)
                modified = True
        if reset:
            titles = self._getDefaultColumn()
        columns = [column.Identifier for column in self._column.getColumns()]
        for name, title in titles.items():
            if name not in columns:
                self._createColumn(name, title)
                modified = True
        if modified:
            self._setDefaultWidths()

# GridColumn private methods
    def _initModel(self, datasource):
        # TODO: ColumnWidth should be assigned after all columns have 
        # TODO: been added to the GridColumn
        widths = self._getSavedWidths(datasource)
        titles = self._getColumnTitles(widths)
        if widths:
            for name in widths:
                if name in titles:
                    title = titles[name]
                    self._createColumn(name, title)
            self._setSavedWidths(widths)
        else:
            for name, title in titles.items():
                self._createColumn(name, title)
            self._setDefaultWidths()

    def _createColumn(self, name, title):
        column = self._column.createColumn()
        column.Title = title
        column.Identifier = name
        index = self._rowset.findColumn(name)
        column.DataColumnIndex = index -1
        self._column.addColumn(column)

    def _setSavedWidths(self, widths):
        for column in self._column.getColumns():
            name = column.Identifier
            flex = len(column.Title)
            column.ColumnWidth = widths[name]
            column.MinWidth = flex * self._factor
            column.Flexibility = 0

    def _setDefaultWidths(self):
        for column in self._column.getColumns():
            flex = len(column.Title)
            width = flex * self._factor
            column.ColumnWidth = width
            column.MinWidth = width
            column.Flexibility = flex

    def _getSavedWidths(self, datasource):
        widths = {}
        if self._name is not None:
            if not self._multi:
                widths = self._widths
            elif datasource in self._widths:
                widths = self._widths[datasource]
        return widths

    def _getColumnWidths(self):
        widths = OrderedDict()
        if self._datasource is not None:
            for column in self._column.getColumns():
                widths[column.Identifier] = column.ColumnWidth
        return widths

    def _getConfigWidths(self):
        config = self._config.getByName(self._name)
        widths = json.loads(config, object_pairs_hook=OrderedDict)
        return widths

    def _getColumnTitles(self, widths):
        if widths:
            titles = self._getWidthTitles(widths)
        else:
            titles = self._getDefaultColumn()
        return titles

    def _getDefaultColumn(self):
        if self._resource is None:
            titles = self._getDefaultTitles(self._max)
        else:
            titles = self._getResourceTitles(self._max)
        return titles

    def _getWidthTitles(self, widths):
        titles = OrderedDict()
        for width in widths:
            if self._resource is None:
                titles[width] = width
            else:
                titles[width] = self._resolver.resolveString(self._resource % width)
        return titles

    def _getDefaultTitles(self, maximum=0):
        titles = OrderedDict()
        metadata = self._rowset.getMetaData()
        for index in range(metadata.getColumnCount()):
            if maximum > 0 and index >= maximum:
                break
            column = metadata.getColumnLabel(index +1)
            titles[column] = column
        return titles

    def _getResourceTitles(self, maximum=0):
        titles = OrderedDict()
        metadata = self._rowset.getMetaData()
        for index in range(metadata.getColumnCount()):
            if maximum > 0 and index >= maximum:
                break
            column = metadata.getColumnLabel(index +1)
            titles[column] = self._resolver.resolveString(self._resource % column)
        return titles
