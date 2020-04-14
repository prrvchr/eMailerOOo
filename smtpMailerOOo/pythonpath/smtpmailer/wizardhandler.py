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
        #self._dbcontext = createService(self.ctx, 'com.sun.star.sdb.DatabaseContext')
        self._listeners = []
        self._disabled = False
        self._address = createService(self.ctx, 'com.sun.star.sdb.RowSet')
        self._address.CommandType = TABLE
        self._address.Filter = self._getFilter(True)
        self._address.ApplyFilter = True
        self._recipient = createService(self.ctx, 'com.sun.star.sdb.RowSet')
        self._recipient.CommandType = QUERY
        self._statement = None
        self._table = None
        self._query = None
        self._database = None
        #self._document = createService(self.ctx, 'com.sun.star.frame.Desktop').CurrentComponent
        self.index = -1

    @property
    def DataSources(self):
        dbcontext = createService(self.ctx, 'com.sun.star.sdb.DatabaseContext')
        return dbcontext.getElementNames()
    @property
    def Connection(self):
        if self._statement is not None:
            return self._statement.getConnection()
        return None
    @property
    def TableNames(self):
        if self.Connection is not None:
            return self.Connection.getTables().getElementNames()
        return ()
    @property
    def ColumnNames(self):
        if self._table is not None:
            return self._table.getColumns().getElementNames()
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
            elif method == 'Dispatch':
                control = event.Source
                handled = self._executeDispatch(control)
                #self._updateControl(window, control)
            elif method == 'Add':
                grid = window.getControl('GridControl2')
                recipients = self._getRecipientFilters()
                rows = window.getControl('GridControl1').getSelectedRows()
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
                handled = self._rowRecipientExecute(recipients)
                self._updateControl(window, grid)
            elif method == 'RemoveAll':
                grid = window.getControl('GridControl2')
                handled = self._rowRecipientExecute([])
                self._updateControl(window, grid)
            self._wizard.updateTravelUI()
            return handled
        except Exception as e:
            print("WizardHandler.callHandlerMethod() ERROR: %s - %s" % (e, traceback.print_exc()))
    def getSupportedMethodNames(self):
        return ('StateChange', 'Dispatch', 'Add', 'AddAll', 'Remove', 'RemoveAll')

    def getDocumentDataSource(self):
        setting = 'com.sun.star.document.Settings'
        document = createService(self.ctx, 'com.sun.star.frame.Desktop').CurrentComponent
        if document.supportsService('com.sun.star.text.TextDocument'):
            return document.createInstance(setting).CurrentDatabaseDataSource
        return ''

    def setDocumentRecord(self, index):
        dispatch = None
        document = createService(self.ctx, 'com.sun.star.frame.Desktop').CurrentComponent
        frame = document.getCurrentController().Frame
        flag = uno.getConstantByName('com.sun.star.frame.FrameSearchFlag.SELF')
        if document.supportsService('com.sun.star.text.TextDocument'):
            url = self._getUrl('.uno:DataSourceBrowser/InsertContent')
            dispatch = frame.queryDispatch(url, '_self', flag)
        elif document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
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
        direct = uno.Enum('com.sun.star.beans.PropertyState', 'DIRECT_VALUE')
        connection = self._recipient.ActiveConnection
        table = connection.getTables().getByIndex(0).Name
        descriptor.append(PropertyValue('DataSourceName', -1, self._recipient.DataSourceName, direct))
        descriptor.append(PropertyValue('ActiveConnection', -1, connection, direct))
        descriptor.append(PropertyValue('Command', -1, table, direct))
        descriptor.append(PropertyValue('CommandType', -1, TABLE, direct))
        descriptor.append(PropertyValue('Cursor', -1, self._recipient, direct))
        descriptor.append(PropertyValue('Selection', -1, (row, ), direct))
        descriptor.append(PropertyValue('BookmarkSelection', -1, False, direct))
        return tuple(descriptor)

    def _updateUI(self, window, event):
        try:
            control = event.Source
            tag = control.Model.Tag
            if tag == 'DataSource':
                datasource = control.SelectedItem
                try:
                    dbcontext = createService(self.ctx, 'com.sun.star.sdb.DatabaseContext')
                    database = dbcontext.getByName(datasource)
                    connection = database.getConnection('', '')
                except WrappedTargetException as e:
                    self._table = None
                    self._database = None
                    self._statement = None
                else:
                    if not database.IsPasswordRequired:
                        self._database = database
                        self._statement = connection.createStatement()
                        self._recipient.DataSourceName = datasource
                        self._address.DataSourceName = datasource
                        self._query = self._getQuery()
                        self._recipient.Command = self._query.Name
                        self._recipient.Filter = self._query.Filter
                        self._recipient.ApplyFilter = True
                        self._recipient.Order = self._query.Order
                        self._address.Order = self._query.Order
                        self._refresh(event)
                    else:
                        self._table = None
                        self._database = None
                        self._statement = None
            elif tag == 'AddressBook':
                print("WizardHandler._updateUI() AddressBook 1")
                table = control.SelectedItem
                self._table = self.Connection.getTables().getByName(table)
                self._query.UpdateTableName = table
                self._address.Command = table
                self._executeRowSet(self._query.Order)
                self._refreshColumns(window.getControl('ListBox2'))
                self._refreshButton(window)
                print("WizardHandler._updateUI() AddressBook 2")
            elif tag == 'Columns':
                print("WizardHandler._updateUI() Columns 1")
                # TODO: During ListBox initializing the listener must be disabled...
                if self._disabled:
                    return True
                # TODO: XRowset.Order should be treated as a column stack
                # TODO: where adding is done at the end and removing will keep order
                orders = self._recipient.Order.strip('"').split('","')
                columns = control.getSelectedItems()
                if len(orders) > len(columns):
                    for order in orders:
                        if order not in columns:
                            orders.remove(order)
                elif len(orders) < len(columns):
                    for column in columns:
                        if column not in orders:
                            orders.append(column)
                order = '"%s"' % '","'.join(orders) if len(orders) else ''
                self._executeRowSet(order)
                print("WizardHandler._updateUI() ************************************************")
                self._refreshButton(window)
                print("WizardHandler._updateUI() Columns 2")
            else:
                pass
                #self._updateControl(window, control)
            return True
        except Exception as e:
            print("WizardHandler._updateUI() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _refreshButton(self, window):
        self._updateControl(window, window.getControl('GridControl1'))
        self._updateControl(window, window.getControl('GridControl2'))

    def _executeRowSet(self, order):
        self._address.Order = order
        self._address.execute()
        self._recipient.Order = self._address.Order
        self._recipient.execute()
        self._query.Order = self._recipient.Order

    def refreshTables(self, control):
        tables = self.TableNames
        control.Model.StringItemList = ()
        control.Model.StringItemList = tables
        table = self._query.UpdateTableName
        table = tables[0] if len(tables) != 0 and table == '' else table
        control.selectItem(table, True)

    def _refreshColumns(self, control):
        self._disabled = True
        control.Model.StringItemList = ()
        control.Model.StringItemList = self.ColumnNames
        columns = self._getOrderIndex()
        control.selectItemsPos(columns, True)
        self._disabled = False

    def _executeDispatch(self, control):
        tag = control.Model.Tag
        document = createService(self.ctx, 'com.sun.star.frame.Desktop').CurrentComponent
        frame = document.CurrentController.Frame
        dispatcher = createService(self.ctx, 'com.sun.star.frame.DispatchHelper')
        if tag == 'DataSource':
            dispatcher.executeDispatch(frame, '.uno:AutoPilotAddressDataSource', '', 0, ())
        elif tag == 'AddressBook':
            dispatcher.executeDispatch(frame, '.uno:AddressBookSource', '', 0, ())
        return True

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

    def _getQuery(self, queryname='smtpMailerOOo'):
        try:
            print("WizardHandler._getQuery() 1")
            queries = self._database.QueryDefinitions
            if queries.hasByName(queryname):
                print("WizardHandler._getQuery() 2")
                query = queries.getByName(queryname)
            else:
                query = createService(self.ctx, 'com.sun.star.sdb.QueryDefinition')
                table = self.Connection.getTables().getByIndex(0)
                column = table.getColumns().getByIndex(0)
                query.Command = 'SELECT * FROM "%s"' % table.Name
                query.UpdateTableName = table.Name
                query.Filter = '0=1'
                query.ApplyFilter = True
                query.Order = '"%s"' % column.Name
                queries.insertByName(queryname, query)
                print("WizardHandler._getQuery() 3")
                self._database.DatabaseDocument.store()
                print("WizardHandler._getQuery() 4")
            return query
        except Exception as e:
            print("WizardHandler._getQuery() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _updateControl(self, window, control):
        try:
            tag = control.Model.Tag
            if tag == 'DataSource':
                window.getControl('ListBox1').Model.StringItemList = self.DataSources
            elif tag == 'Addresses':
                selected = control.hasSelectedRows()
                enabled = control.Model.GridDataModel.RowCount != 0
                window.getControl('CommandButton1').Model.Enabled = enabled
                window.getControl('CommandButton2').Model.Enabled = selected
            elif tag == 'Recipients':
                selected = control.hasSelectedRows()
                enabled = control.Model.GridDataModel.RowCount != 0
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

    def _getOrderIndex(self):
        index = []
        orders = self._query.Order.strip('"').split('","')
        columns = self.ColumnNames
        for order in orders:
            if order in columns:
                index.append(columns.index(order))
        return tuple(index)

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
