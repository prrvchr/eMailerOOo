#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.util import XRefreshable
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.beans import PropertyValue
from com.sun.star.util import URL
from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE
from com.sun.star.lang import WrappedTargetException

from unolib import PropertySet
from unolib import createService
from unolib import getProperty
from unolib import getStringResource

from .griddatamodel import GridDataModel
from .dbqueries import getSqlQuery

from .configuration import g_identifier
from .configuration import g_extension
from .configuration import g_column_index
from .configuration import g_column_filters
from .configuration import g_table_index

from .logger import logMessage

import traceback


class WizardHandler(unohelper.Base,
                    PropertySet,
                    XContainerWindowEventHandler,
                    XRefreshable):
    def __init__(self, ctx, wizard):
        self.ctx = ctx
        self._wizard = wizard
        self._dbcontext = createService(self.ctx, 'com.sun.star.sdb.DatabaseContext')
        self._table = None
        self._columns = []
        self._statement = None
        self._listeners = []
        self._address = createService(self.ctx, 'com.sun.star.sdb.RowSet')
        self._address.CommandType = TABLE
        self._address.Filter = self._getFilter(True)
        self._address.ApplyFilter = True
        self._recipient = createService(self.ctx, 'com.sun.star.sdb.RowSet')
        self._recipient.CommandType = QUERY
        self._query = None
        self._database = None
        self._document = createService(self.ctx, 'com.sun.star.frame.Desktop').CurrentComponent
        self.index = -1

    @property
    def DataSources(self):
        return self._dbcontext.ElementNames
    @property
    def Connection(self):
        if self._statement is not None:
            return self._statement.getConnection()
        return None
    @property
    def TableNames(self):
        if self._statement is not None:
            return self._statement.getConnection().getTables().ElementNames
        return ()
    @property
    def ColumnNames(self):
        if self._statement is not None:
            if self._table is not None:
                return self._table.Columns.ElementNames
        return ()

    # XRefreshable
    def refresh(self):
        pass
    def addRefreshListener(self, listener):
        if listener not in self._listeners:
            self._listeners.append(listener)
    def removeRefreshListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)
    def _refresh(self, event):
        for listener in self._listeners:
            listener.refreshed(event)

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        try:
            handled = False
            if method == 'StateChange':
                handled = self._updateUI(window, event)
            elif method == 'Add':
                grid = window.getControl('GridControl2')
                recipients = self._getRecipientFilters()
                rows = window.getControl('GridControl1').getSelectedRows()
                print("wizardhandler.callHandlerMethod() Add %s - %s)" % (recipients, rows))
                filters = self._getAddressFilters(rows, recipients)
                handled = self._rowRecipientExecute(recipients + filters)
                self._updateControl(window, grid)
            elif method == 'AddAll':
                grid = window.getControl('GridControl2')
                recipients = self._getRecipientFilters()
                rows = range(self._address.RowCount)
                filters = self._getAddressFilters(rows, recipients)
                handled = self._rowRecipientExecute(recipients + filters)
                self._updateControl(window, grid)
            elif method == 'Remove':
                grid = window.getControl('GridControl2')
                rows = grid.getSelectedRows()
                recipients = self._getRecipientFilters(rows)
                print("wizardhandler.callHandlerMethod() Remove %s - %s)" % (recipients, rows))
                handled = self._rowRecipientExecute(recipients)
                self._updateControl(window, grid)
            elif method == 'RemoveAll':
                handled = self._rowRecipientExecute([])
                self._updateControl(window, window.getControl('GridControl2'))
            self._wizard.updateTravelUI()
            return handled
        except Exception as e:
            print("WizardHandler.callHandlerMethod() ERROR: %s - %s" % (e, traceback.print_exc()))
    def getSupportedMethodNames(self):
        return ('StateChange', 'Add', 'AddAll', 'Remove', 'RemoveAll')

    def getDocumentDataSource(self):
        setting = 'com.sun.star.document.Settings'
        if self._document.supportsService('com.sun.star.text.TextDocument'):
            return self._document.createInstance(setting).CurrentDatabaseDataSource
        return ''

    def setDocumentRecord(self, index):
        dispatch = None
        frame = self._document.CurrentController.Frame
        flag = uno.getConstantByName('com.sun.star.frame.FrameSearchFlag.SELF')
        if self._document.supportsService('com.sun.star.text.TextDocument'):
            url = self._getUrl('.uno:DataSourceBrowser/InsertContent')
            dispatch = frame.queryDispatch(url, '_self', flag)
        elif self._document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
            url = self._getUrl('.uno:DataSourceBrowser/InsertColumns')
            dispatch = frame.queryDispatch(url, '_self', flag)
        if dispatch is not None:
            dispatch.dispatch(url, self._getDataDescriptor(index + 1))
        self.index = index

    def _getUrl(self, complete):
        url = URL()
        url.Complete = complete
        dummy, url = createService(self.ctx, 'com.sun.star.util.URLTransformer').parseStrict(url)
        return url

    def _getDataDescriptor(self, row):
        descriptor = []
        value = uno.Enum('com.sun.star.beans.PropertyState', 'DIRECT_VALUE')
        connection = self._recipient.ActiveConnection
        table = connection.getTables().getByIndex(0).Name
        descriptor.append(PropertyValue('DataSourceName', -1, self._recipient.DataSourceName, value))
        descriptor.append(PropertyValue('ActiveConnection', -1, connection, value))
        descriptor.append(PropertyValue('Command', -1, table, value))
        descriptor.append(PropertyValue('CommandType', -1, TABLE, value))
        descriptor.append(PropertyValue('Cursor', -1, self._recipient, value))
        descriptor.append(PropertyValue('Selection', -1, [row], value))
        descriptor.append(PropertyValue('BookmarkSelection', -1, False, value))
        return descriptor

    def _getColumnModel(self, rows=()):
        if rows is None:
            rows = range(2)
        service = 'com.sun.star.awt.grid.DefaultGridColumnModel'
        model = self.ctx.ServiceManager.createInstance(service)
        i = 0
        for name in self.ColumnNames:
            if i in rows:
                column = model.createColumn()
                column.Title = name
                column.MinWidth = 4 * len(name)
                column.DataColumnIndex = i
                model.addColumn(column)
            i += 1
        return model

    def _updateUI(self, window, event):
        try:
            control = event.Source
            tag = control.Model.Tag
            if tag == 'DataSource':
                datasource = control.SelectedItem
                try:
                    database = self._dbcontext.getByName(datasource)
                except WrappedTargetException as e:
                    self._statement = None
                    self._table = None
                else:
                    if not database.IsPasswordRequired:
                        connection = database.getConnection('', '')
                        self._database = database
                        self._query = self._getQuery(connection)
                        self._statement = connection.createStatement()
                        self._address.DataSourceName = datasource
                        self._address.Order = self._query.Order
                        self._recipient.DataSourceName = datasource
                        self._recipient.Command = self._query.Name
                        self._recipient.Filter = self._query.Filter
                        self._recipient.ApplyFilter = True
                        self._recipient.Order = self._query.Order
                    else:
                        self._query = None
                        self._statement = None
                        self._table = None
                        self._database = None
                self._refresh(event)
            elif tag == 'AddressBook':
                self._table = self.Connection.getTables().getByName(control.SelectedItem)
                self._query.UpdateTableName = self._table.Name
                self._address.Command = control.SelectedItem
                window.getControl('ListBox2').Model.StringItemList = self.ColumnNames
                self._address.execute()
                self._updateControl(window, window.getControl('GridControl1'))
            elif tag == 'Columns':
                index = control.getSelectedItemsPos()
                self._columns = index
                self._setQueryOrder(control.getSelectedItems())
                grid1 = window.getControl('GridControl1')
                grid2 = window.getControl('GridControl2')
                grid1.Model.ColumnModel = self._getColumnModel(index)
                grid2.Model.ColumnModel = self._getColumnModel(index)
                self._address.execute()
                self._recipient.execute()
                self._updateControl(window, grid1)
                self._updateControl(window, grid2)
            else:
                pass
                #self._updateControl(window, control)
            return True
        except Exception as e:
            print("WizardHandler._updateUI() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _setQueryOrder(self, columns):
        order = ''
        if len(columns):
            order = '"%s"' % '","'.join(columns)
        self._query.Order = order
        self._address.Order = order
        self._recipient.Order = order

    def _rowRecipientExecute(self, filters=None):
        self._query.ApplyFilter = False
        self._recipient.ApplyFilter = False
        if filters is not None:
            self._query.Filter = self._getQueryFilter(filters)
        self._recipient.Filter = self._query.Filter
        self._recipient.ApplyFilter = True
        self._query.ApplyFilter = True
        self._recipient.execute()
        return True

    def _getQuery(self, connection, queryname="smtpMailerOOo"):
        queries = self._database.QueryDefinitions
        if queries.hasByName(queryname):
            query = queries.getByName(queryname)
        else:
            query = self.ctx.ServiceManager.createInstance("com.sun.star.sdb.QueryDefinition")
            table = connection.getTables().getByIndex(0)
            column = table.getColumns().getByIndex(0)
            query.Command = 'SELECT * FROM "%s"' % table.Name
            query.UpdateTableName = table.Name
            query.Filter = '0=1'
            query.ApplyFilter = True
            query.Order = '"%s"' % column.Name
            queries.insertByName(queryname, query)
            self._database.DatabaseDocument.store()
        return query

    def _updateControl(self, window, control):
        try:
            tag = control.Model.Tag
            selected = control.hasSelectedRows()
            enabled = control.Model.GridDataModel.RowCount != 0
            if tag == 'Addresses':
                window.getControl('CommandButton1').Model.Enabled = enabled
                window.getControl('CommandButton2').Model.Enabled = selected
            elif tag == 'Recipients':
                window.getControl('CommandButton3').Model.Enabled = selected
                window.getControl('CommandButton4').Model.Enabled = enabled
        except Exception as e:
            print("WizardHandler._updateControl() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _getQueryFilter(self, filters=None):
        if filters is None:
            result = ''
        else:
            result = "%s IN ('%s')" % ('"Resource"', "','".join(filters))
        return result

    def _getFilter(self, any=False):
        filters = []
        for column in g_column_filters:
            filters.append('"%s" IS NOT NULL' % column)
        separator = " OR " if any else " AND "
        return separator.join(filters)

    def _getColumn(self, index=None):
        index = g_column_index if index is None else index
        return self._table.Columns.getByIndex(index)

    def _getColumnIndex(self, name):
        if self._table is not None and self._table.Columns.hasByName(name):
            column = self._table.Columns.getByName(name)
            mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
            mri.inspect(column)
            return column
        return -1

    def getOrderIndex(self):
        index = []
        orders = self._query.Order.strip('"').split('","')
        columns = self.ColumnNames
        print("wizardhandler.getOrderIndex() %s - %s" % (orders, columns))
        for order in orders:
            if order in columns:
                index.append(columns.index(order))
        return index

    def _getRecipientRowNum(self, index):
        rownum = -1
        column = self._recipient.getColumns().getCount()
        if self._recipient.absolute(index +1):
            rownum = self._recipient.getLong(column)
        return rownum

    def _getRecipientFilters(self, rows=(), index=0):
        result = []
        self._recipient.beforeFirst()
        while self._recipient.next():
            row = self._recipient.Row -1
            if row not in rows:
                result.append(self._recipient.getString(index +1))
        print("wizardhandler._getRecipientFilters() %s - %s)" % (self._recipient.RowCount, result))
        return result

    def _getAddressFilters(self, rows, filters=(), index=0):
        result = []
        for row in rows:
            self._address.absolute(row +1)
            filter = self._address.getString(index +1)
            if filter not in filters:
                result.append(filter)
        print("wizardhandler._getAddressFilters() %s - %s)" % (self._address.RowCount, result))
        return result

    def _getPropertySetInfo(self):
        properties = {}
        readonly = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.READONLY')
        transient = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.TRANSIENT')
        properties['Connection'] = getProperty('Connection', 'com.sun.star.sdbc.XConnection', transient)
        properties['DataSources'] = getProperty('DataSources', '[] string', transient)
        properties['TableNames'] = getProperty('TableNames', '[] string', transient)
        properties['ColumnNames'] = getProperty('ColumnNames', '[] string', transient)
        return properties
