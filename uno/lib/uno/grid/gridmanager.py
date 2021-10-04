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

from ..unotool import createService
from ..unotool import getConfiguration
from ..unotool import getResourceLocation
from ..unotool import getStringResource

from ..dbtool import isSimilar

from ..configuration import g_extension
from ..configuration import g_identifier

from .griddata import GridData
from .gridview import GridView

import json
from threading import Thread
from collections import OrderedDict
import traceback


class GridManager(unohelper.Base):
    def __init__(self, ctx, rowset, parent, possize, config, resource=None, maxi=None, multi=False, name='Grid1'):
        self._ctx = ctx
        self._factor = 5
        self._datasource = None
        self._query = None
        self._resource = resource
        if resource is not None:
            self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._name = config
        self._widths = {}
        if config is not None:
            self._config = getConfiguration(ctx, g_identifier, True)
            self._widths = self._getConfigWidths()
        self._max = maxi
        self._multi = multi
        self._similar = True
        self._rowset = rowset
        self._url = getResourceLocation(ctx, g_identifier, g_extension)
        service = 'com.sun.star.awt.grid.DefaultGridColumnModel'
        self._model = createService(ctx, service)
        self._grid = GridData()
        self._columns = {}
        self._composer = None
        self._view = GridView(ctx, name, self, parent, possize)

# GridManager getter methods
    def getGridModels(self):
        return self._grid, self._model

    def getSelectedRows(self):
        return self._view.getSelectedRows()

# GridManager setter methods
    def dispose(self):
        self._model.dispose()
        self._grid.dispose()

    def addSelectionListener(self, listener):
        self._view.getGrid().addSelectionListener(listener)

    def removeSelectionListener(self, listener):
        self._view.getGrid().removeSelectionListener(listener)

    def showControls(self, state):
        self._view.setWindowPosSize(state)

    def resetGrid(self):
        self._view.showGridColumnHeader(False)
        self._grid.resetRowSetData()

    def setRowSetData(self, rowset):
        name = self._view.getGrid().Model.Name
        connection = rowset.ActiveConnection
        datasource = connection.Parent
        query = rowset.UpdateTableName
        if self._isDataSourceChanged(datasource, query):
            print("GridManager.setRowSetData() 1 %s" % datasource.Name)
            if self._datasource is not None:
                self.saveWidths()
            self._composer = self._getComposer(connection, rowset)
            if self._multi:
                self._similar = isSimilar(connection)
            self._columns = self._getColumns(rowset.getMetaData())
            identifiers = self._initColumnModel(datasource)
            self._initColumns(identifiers)
            self._datasource = datasource
            self._query = query
            self._view.showGridColumnHeader(True)
        self._grid.setRowSetData(rowset)

    def saveWidths(self):
        width = self._getColumnWidths()
        if not self._multi:
            self._widths = width
        elif self._datasource is not None:
            name = self._getDataSourceName(self._datasource)
            self._widths[name] = width

    def saveColumnWidths(self):
        self.saveWidths()
        columns = json.dumps(self._widths)
        name = self._getConfigWidthName()
        self._config.replaceByName(name, columns)
        self._config.commitChanges()

    def setColumn(self, identifier, add, reset, index):
        self._view.deselectColumn(index)
        if reset:
            modified, identifiers = self._resetColumn()
        else:
            identifiers = [column.Identifier for column in self._model.getColumns()]
            if add:
                modified = self._addColumn(identifiers, identifier)
            else:
                modified = self._removeColumn(identifiers, identifier)
        if modified:
            self._setDefaultWidths()
            self._view.setColumns(self._url, identifiers)

    def isSelected(self, image):
        return image.endswith(self._view.getSelected())

    def isUnSelected(self, image):
        return image.endswith(self._view.getUnSelected())

    def setOrder(self, identifier, add, index):
        self._setOrder(identifier, add, index)
        Thread(target=self._executeRowSet).start()

# GridManager private methods
    def _isDataSourceChanged(self, datasource, query):
        return self._datasource != datasource or self._query != query

    def _getDataSourceName(self, datasource):
        if self._similar:
            name = datasource.Name
        else:
            table = self._composer.getTables().getByIndex(0).Name
            name = '%s.%s' % (datasource.Name, table)
        return name

    def _resetColumn(self):
        self._removeColumns()
        identifiers = self._getDefaultIdentifiers()
        for identifier in identifiers:
            self._createColumn(identifier)
        return True, identifiers

    def _addColumn(self, identifiers, identifier):
        modified = False
        if identifier not in identifiers:
            if self._createColumn(identifier):
                identifiers.append(identifier)
                modified = True
        return modified

    def _removeColumn(self, identifiers, identifier):
        modified = False
        if identifier in identifiers:
            if self._removeIdentifier(identifier):
                identifiers.remove(identifier)
                modified = True
        return modified

    def _setOrder(self, identifier, add, index):
        self._view.deselectOrder(index)
        if add:
            self._addOrder(identifier)
        else:
            self._removeOrder(identifier)
        order = self._composer.getOrder()
        self._rowset.Order = order
        if self._query:
            self._saveQueryOrder(order)
        else:
            self._saveConfigOrder(order)
        self._view.setOrders(self._url, self._getOrders())

    def _saveQueryOrder(self, order):
        query = self._datasource.getQueryDefinitions().getByName(self._query)
        self._composer.setQuery(query.Command)
        self._composer.setOrder(order)
        query.Command = self._composer.getQuery()
        self._datasource.DatabaseDocument.store()

    def _saveConfigOrder(self, order):
        name = self._getConfigOrderName()
        self._config.replaceByName(name, self._rowset.Order)
        self._config.commitChanges()

    def _executeRowSet(self):
        self._rowset.execute()

    def _addOrder(self, identifier):
        ascending = self._view.getSortDirection()
        column = self._composer.getColumns().getByName(identifier)
        self._composer.appendOrderByColumn(column, ascending)

    def _removeOrder(self, identifier):
        orders = []
        enumeration = self._composer.getOrderColumns().createEnumeration()
        while enumeration.hasMoreElements():
            column = enumeration.nextElement()
            if column.Name != identifier:
                orders.append(column)
        self._composer.Order = ''
        for order in orders:
            self._composer.appendOrderByColumn(order, order.IsAscending)

    def _getOrders(self):
        orders = OrderedDict()
        enumeration = self._composer.getOrderColumns().createEnumeration()
        while enumeration.hasMoreElements():
            column = enumeration.nextElement()
            orders[column.Name] = column.IsAscending
        return orders

    def _getColumns(self, metadata):
        columns = OrderedDict()
        for index in range(metadata.getColumnCount()):
            column = metadata.getColumnLabel(index +1)
            columns[column] = self._getColumnTitle(column)
        return columns

    def _getComposer(self, connection, rowset):
        name = self._view.getGrid().Model.Name
        composer = connection.getComposer(rowset.CommandType, rowset.Command)
        composer.setCommand(rowset.Command, rowset.CommandType)
        composer.setOrder(rowset.Order)
        return composer

    def _initColumnModel(self, datasource):
        # TODO: ColumnWidth should be assigned after all columns have 
        # TODO: been added to the GridColumnModel
        self._removeColumns()
        widths = self._getSavedWidths(datasource)
        identifiers = self._getIdentifiers(widths)
        name = self._view.getGrid().Model.Name
        if widths:
            for identifier in widths:
                self._createColumn(identifier)
            self._setSavedWidths(widths)
        else:
            for identifier in identifiers:
                self._createColumn(identifier)
            self._setDefaultWidths()
        return identifiers

    def _initColumns(self, identifiers):
        self._view.initColumns(self._url, self._columns, identifiers)
        orders = self._composer.getOrderColumns().createEnumeration()
        self._view.initOrders(self._url, self._columns, orders)

    def _getSavedWidths(self, datasource):
        widths = {}
        if self._multi:
            name = self._getDataSourceName(datasource)
            if name in self._widths:
                widths = self._widths[name]
        else:
            widths = self._widths
        return widths

    def _getIdentifiers(self, widths):
        identifiers = []
        for identifier in widths:
            if identifier in self._columns:
                identifiers.append(identifier)
        if not identifiers:
            identifiers = self._getDefaultIdentifiers()
        return identifiers

    def _removeColumns(self):
        for index in range(self._model.getColumnCount() -1, -1, -1):
            self._model.removeColumn(index)

    def _createColumn(self, identifier):
        created = False
        if identifier in self._columns:
            column = self._model.createColumn()
            column.Identifier = identifier
            column.Title = self._columns[identifier]
            indexes = tuple(self._columns.keys())
            column.DataColumnIndex = indexes.index(identifier)
            self._model.addColumn(column)
            created = True
        return created

    def _removeIdentifier(self, identifier):
        removed = False
        for index in range(self._model.getColumnCount() -1, -1, -1):
            column = self._model.getColumn(index)
            if column.Identifier == identifier:
                self._model.removeColumn(index)
                removed = True
                break
        return removed

    def _setSavedWidths(self, widths):
        for column in self._model.getColumns():
            identifier = column.Identifier
            flex = len(column.Title)
            column.MinWidth = flex * self._factor
            column.Flexibility = 0
            if identifier in widths:
                column.ColumnWidth = widths[identifier]
            else:
                column.ColumnWidth = flex * self._factor

    def _setDefaultWidths(self):
        for column in self._model.getColumns():
            flex = len(column.Title)
            width = flex * self._factor
            column.ColumnWidth = width
            column.MinWidth = width
            column.Flexibility = flex

    def _getColumnWidths(self):
        widths = OrderedDict()
        if self._datasource is not None:
            for column in self._model.getColumns():
                widths[column.Identifier] = column.ColumnWidth
        return widths

    def _getConfigWidths(self):
        name = self._getConfigWidthName()
        config = self._config.getByName(name)
        widths = json.loads(config, object_pairs_hook=OrderedDict)
        return widths

    def _getDefaultIdentifiers(self):
        identifiers = tuple(self._columns.keys())
        return identifiers[slice(self._max)]

    def _getColumnTitle(self, identifier):
        if self._resource is None:
            title = identifier
        else:
            title = self._resolver.resolveString(self._resource % identifier)
        return title

    def _getConfigWidthName(self):
        return '%s%s' % (self._name, 'Columns')

    def _getConfigOrderName(self):
        return '%s%s' % (self._name, 'Orders')
