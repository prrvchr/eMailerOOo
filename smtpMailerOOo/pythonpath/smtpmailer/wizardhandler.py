#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.util import XRefreshable
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE
from com.sun.star.lang import WrappedTargetException

from unolib import PropertySet
from unolib import createService
from unolib import getStringResource

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
        self._address = self.ctx.ServiceManager.createInstance('com.sun.star.sdb.RowSet')
        self._address.CommandType = TABLE
        self._recipient = self.ctx.ServiceManager.createInstance('com.sun.star.sdb.RowSet')
        self._recipient.CommandType = QUERY

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
            return handled
        except Exception as e:
            print("WizardHandler.callHandlerMethod() ERROR: %s - %s" % (e, traceback.print_exc()))
    def getSupportedMethodNames(self):
        return ('StateChange', 'Add', 'AddAll', 'Remove', 'RemoveAll')

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
                        #mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
                        #mri.inspect(connection)
                        self._statement = connection.createStatement()
                        #self._table = connection.getTables().getByIndex(g_table_index)
                        self._address.ActiveConnection = connection
                        self._recipient.ActiveConnection = connection
                    else:
                        self._statement = None
                        self._table = None
                self._refresh(event)
            elif tag == 'AddressBook':
                #mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
                #mri.inspect(self.Connection)
                self._table = self.Connection.getTables().getByName(control.SelectedItem)
                self._address.Command = control.SelectedItem
                control = window.getControl('ListBox2')
                control.Model.StringItemList = self.ColumnNames
                #self._address.Filter = self._getQueryFilter()
                #self._address.ApplyFilter = True
                #self._address.Order = self._getQueryOrder()
                self._address.execute()
                #self._table = self._address.ActiveConnection.Tables.getByIndex(self._tableindex)
                #connection = database.getConnection("","")
                #mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
                #mri.inspect(connection)
                #self._controller.Connection = connection
            elif tag == 'Columns':
                #mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
                #mri.inspect(control)
                self._columns = control.SelectedItemsPos
                self._address.execute()
            self._wizard.updateTravelUI()
            return True
        except Exception as e:
            print("WizardHandler._updateUI() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _getQueryFilter(self, filters=None):
        if filters is None:
            result = "(%s)" % (self._getFilter())
        elif len(filters):
            result = "(%s)" % (" OR ".join(filters))
        else:
            result = "(%s AND \"%s\" = '')" % (self._getFilter(), self._getColumn().Name)
        return result

    def _getFilter(self):
        filter = []
        for column in g_column_filters:
            filter.append('"%s" IS NOT NULL' % (self._getColumn(column).Name))
        return " AND ".join(filter)

    def _getColumn(self, index=None):
        index = g_column_index if index is None else index
        return self._table.Columns.getByIndex(index)

    def _getQueryOrder(self):
        order = '"%s"' % (self._getColumn().Name)
        return order

    def _getPropertySetInfo(self):
        properties = {}
        readonly = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.READONLY')
        transient = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.TRANSIENT')
        properties['Connection'] = getProperty('Connection', 'com.sun.star.sdbc.XConnection', transient)
        properties['DataSources'] = getProperty('DataSources', '[] string', transient)
        properties['TableNames'] = getProperty('TableNames', '[] string', transient)
        properties['ColumnNames'] = getProperty('ColumnNames', '[] string', transient)
        return properties
