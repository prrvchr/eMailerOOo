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

import uno
import unohelper

from com.sun.star.beans import PropertyValue

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.view.SelectionType import MULTI

from com.sun.star.uno import Exception as UnoException

from com.sun.star.document.MacroExecMode import ALWAYS_EXECUTE_NO_WARN

from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.sdb.SQLFilterOperator import EQUAL
from com.sun.star.sdb.SQLFilterOperator import SQLNULL
from com.sun.star.sdb.SQLFilterOperator import NOT_SQLNULL

from .mergerhandler import DispatchListener

from ..grid import GridManager
from ..grid import GridModel

from ..mail import MailModel

from ..unotool import createService
from ..unotool import executeDispatch
from ..unotool import executeFrameDispatch
from ..unotool import getConfiguration
from ..unotool import getDesktop
from ..unotool import getInteractionHandler
from ..unotool import getPathSettings
from ..unotool import getPropertyValue
from ..unotool import getPropertyValueSet
from ..unotool import getResourceLocation
from ..unotool import getSimpleFile
from ..unotool import getStringResource
from ..unotool import getUrl
from ..unotool import getUrlPresentation

from ..dbtool import getResultValue
from ..dbtool import getSequenceFromResult
from ..dbtool import getValueFromResult

from ..dbqueries import getSqlQuery

from ..mailertool import getDocument

from ..logger import getMessage
from ..logger import logMessage

from ..configuration import g_identifier
from ..configuration import g_extension
from ..configuration import g_fetchsize

from collections import OrderedDict
from six import string_types
from threading import Thread
from threading import Condition
from time import sleep
import json
import traceback


class MergerModel(MailModel):
    def __init__(self, ctx, datasource):
        self._ctx = ctx
        self._datasource = datasource
        self._configuration = getConfiguration(ctx, g_identifier, True)
        self._document = getDesktop(ctx).CurrentComponent
        #mri = createService(self._ctx, 'mytools.Mri')
        #mri.inspect(self._document)
        self._path = self._getPath()
        self._prefix = 'Sub.'
        service = 'com.sun.star.sdb.DatabaseContext'
        self._dbcontext = createService(ctx, service)
        self._addressbook = None
        self._statement = None
        self._queries = None
        self._address = self._getRowSet(TABLE)
        self._recipient = self._getRowSet(QUERY)
        self._grid1 = None
        self._grid2 = None
        self._composer = None
        self._subcomposer = None
        self._name = None
        self._rows = ()
        self._tables = {}
        self._query = None
        self._subquery = None
        self._identifiers = ()
        self._emails = ()
        self._changed = False
        self._disposed = False
        self._similar = False
        self._temp = False
        self._saved = False
        self._lock = Condition()
        self._url = getResourceLocation(ctx, g_identifier, g_extension)
        self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._resources = {'Step': 'MergerPage%s.Step',
                           'Title': 'MergerPage%s.Title',
                           'TabTitle': 'MergerTab%s.Title',
                           'Progress': 'MergerPage1.Label6.Label.%s',
                           'Error': 'MergerPage1.Label8.Label.%s',
                           'Index': 'MergerPage1.Label14.Label.%s',
                           'Message': 'MergerTab2.Label1.Label',
                           'Recipient': 'MailWindow.Label4.Label',
                           'PickerTitle': 'Mail.FilePicker.Title',
                           'Property': 'Mail.Document.Property.%s',
                           'Document': 'MailWindow.Label8.Label.1'}

    @property
    def Connection(self):
        return self._statement.getConnection()

    def _isConnectionNotClosed(self):
        return self._statement is not None

    def _closeConnection(self):
        connection = self._statement.getConnection()
        #self._statement.close()
        self._statement.dispose()
        self._statement = None
        #self._address.dispose()
        #self._address = None
        #self._recipient.dispose()
        #self._recipient = None
        self._composer.dispose()
        self._composer = None
        self._subcomposer.dispose()
        self._subcomposer = None
        self._queries.dispose()
        self._queries = None
        self._addressbook.DatabaseDocument.dispose()
        self._addressbook.dispose()
        self._addressbook = None
        connection.close()

# Procedures called by WizardController
    def dispose(self):
        print("MergerModel.dispose() 1")
        if self._isConnectionNotClosed():
            self._closeConnection()
        self._datasource.dispose()
        if self._grid1 is not None:
            self._grid1.dispose()
        if self._grid2 is not None:
            self._grid2.dispose()
        self._disposed = True
        print("MergerModel.dispose() 2")

    def saveGrids(self):
        print("MergerModel.save() 1")
        if self._grid1 is not None:
            self._grid1.saveColumnSettings()
        if self._grid2 is not None:
            self._grid2.saveColumnSettings()
        print("MergerModel.save() 2")

# Procedures called by WizardPage1
    # AddressBook methods
    def getAvailableAddressBooks(self):
        return self._dbcontext.getElementNames()

    def getDefaultAddressBook(self):
        addressbook = ''
        if self._loadAddressBook():
            addressbook = self._getDocumentAddressBook(addressbook)
        return addressbook

    def _loadAddressBook(self):
        return self._configuration.getByName('MergerLoadDataSource')

    def _getDocumentAddressBook(self, addressbook):
        service = 'com.sun.star.text.TextDocument'
        if self._document.supportsService(service):
            service = 'com.sun.star.document.Settings'
            settings = self._document.createInstance(service)
            addressbook = settings.CurrentDatabaseDataSource
        return addressbook

    def isAddressBookNotLoaded(self, addressbook):
        return self._addressbook is None or self._addressbook.Name != addressbook

    def setAddressBook(self, *args):
        Thread(target=self._setAddressBook, args=args).start()

    # AddressBook private methods
    def _setAddressBook(self, addressbook, progress, setAddressBook):
        try:
            step = 2
            sleep(0.2)
            progress(5)
            queries = label = message = None
            self._tables = {}
            progress(10)
            # FIXME: If changes have been made then save them...
            if self._queries is not None:
                self._saveQueries()
            # FIXME: If an addressbook has been loaded we need:
            # FIXME: to dispose all components who use the connection and close the connection
            if self._isConnectionNotClosed():
                self._closeConnection()
            progress(20)
            try:
                datasource = self._getDataSource(addressbook)
                progress(30)
                if datasource.IsPasswordRequired:
                    handler = getInteractionHandler(self._ctx)
                    connection = datasource.getIsolatedConnectionWithCompletion(handler)
                else:
                    connection = datasource.getIsolatedConnection('', '')
                progress(40)
                service = 'com.sun.star.sdb.SingleSelectQueryComposer'
                if service not in connection.getAvailableServiceNames():
                    msg = self._getErrorMessage(2, service)
                    e = self._getUnoException(msg)
                    raise e
                #if not connection.getMetaData().supportsCorrelatedSubqueries():
                #    msg = self._getErrorMessage(3)
                #    e = self._getUnoException(msg)
                #    raise e
            except UnoException as e:
                format = (addressbook, e.Message)
                message = self._getErrorMessage(0, format)
            else:
                progress(50)
                #mri = createService(self._ctx, 'mytools.Mri')
                #mri.inspect(connection)
                self._addressbook = datasource
                self._statement = connection.createStatement()
                self._composer = connection.createInstance(service)
                self._subcomposer = connection.createInstance(service)
                progress(60)
                self._tables, self._similar = self._getTablesInfos(connection)
                progress(70)
                self._address.ActiveConnection = connection
                self._recipient.ActiveConnection = connection
                self._queries = datasource.getQueryDefinitions()
                progress(80)
                composer = connection.createInstance(service)
                queries = self._getQueries(composer)
                progress(90)
                self._setSubQueryTable(composer, queries)
                composer.dispose()
                label = self._getIndexLabel()
                step = 3
            progress(100)
            setAddressBook(step, queries, self._getTables(), label, message)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def _saveQueries(self):
        saved = False
        if self._query is not None:
            saved |= self._saveQuery(self._query, self._composer)
        if self._subquery is not None:
            saved |= self._saveQuery(self._subquery.First, self._subcomposer)
        if saved:
            self._addressbook.DatabaseDocument.store()

    def _getTablesInfos(self, connection):
        infos = {}
        similar = True
        tables = connection.getTables()
        if tables.hasElements():
            columns = tables.getByIndex(0).getColumns().getElementNames()
            for name in tables.getElementNames():
                info = tables.getByName(name).getColumns().getElementNames()
                if columns != info:
                    similar = False
                infos[name] = info
        return infos, similar

    def _getTables(self):
        return tuple(self._tables.keys())

    def _getDataSource(self, addressbook):
        # We need to check if the registered datasource has an existing odb file
        sf = getSimpleFile(self._ctx)
        location = self._dbcontext.getDatabaseLocation(addressbook)
        if not sf.exists(location):
            msg = self._getErrorMessage(1, location)
            e = self._getUnoException(msg)
            raise e
        if self._temp:
            datasource = self._getTempDataSource(sf, addressbook, location)
        else:
            datasource = self._dbcontext.getByName(addressbook)
        return datasource

    def _getTempDataSource(self, sf, addressbook, location):
        # FIXME: We can undo all changes if the wizard is canceled
        # FIXME: or abort the Wizard while keeping the work already done
        # FIXME: The wizard must be modified to take into account the Cancel button
        url = self._getTempUrl(addressbook)
        if not sf.exists(url):
            sf.copy(location, url)
        datasource = self._dbcontext.getByName(url)
        return datasource

    def _getUnoException(self, msg):
        e = UnoException()
        e.Message = msg
        e.Context = self
        return e

    def _getTempUrl(self, addressbook):
        temp = getPathSettings(self._ctx).Temp
        url = '%s/%s.odb' % (temp, addressbook)
        return url

    def _getQueries(self, composer):
        queries = {}
        names = self._queries.getElementNames()
        for name in names:
            table = self._getQueryTable(composer, names, name)
            if table is not None:
                queries[name] = uno.createUnoStruct('com.sun.star.beans.StringPair', table, '')
        return queries

    def _getQueryTable(self, composer, queries, query):
        table = None
        composer.setCommand(query, QUERY)
        for name in composer.getTables().getElementNames():
            if name in queries:
                table = name
        print("MergerModel._getQueryTable() 1 Table: '%s'" % (table, ))
        return table

    def _setSubQueryTable(self, composer, queries):
        for subquery in queries.values():
            if self._queries.hasByName(subquery.First):
                command = self._queries.getByName(subquery.First).Command
                subquery.Second = self._getSubQueryTable(composer, command)

    def _getSubQueryTable(self, composer, command):
        table = None
        composer.setQuery(command)
        tables = composer.getTables()
        print("MergerModel._getSubQueryTable() Tables.getCount() : %s" % tables.getCount())
        if tables.getCount() > 0:
            table = tables.getElementNames()[0]
        print("MergerModel._getSubQueryTable() Table : %s" % table)
        return table

    def _checkQueries(self, query, subquery):
        return self._checkQuery(query, subquery) and self._checkSubQuery(subquery)
 
    def _checkQuery(self, query, subquery):
        self._composer.setCommand(query, QUERY)
        tables = self._composer.getTables()
        return tables.hasElements() and tables.getByIndex(0).Name == subquery

    def _checkSubQuery(self, subquery):
        self._subcomposer.setCommand(subquery, QUERY)
        tables = self._subcomposer.getTables()
        return tables.hasElements()

    def _getSubQueries(self, queries):
        subqueries = {}
        for query in queries:
            self._composer.setCommand(query, QUERY)
            tables = self._composer.getTables()
            if tables.hasElements() and tables.getElementNames()[0] in queries:
                subqueries[query] = tables.getElementNames()[0]
        return subqueries

    # AddressBook Table methods
    def setAddressBookTable(self, name):
        columns = ()
        tables = self.Connection.getTables()
        if tables.hasByName(name):
            columns = tables.getByName(name).getColumns().getElementNames()
        return columns

    # Query methods
    def isQueryValid(self, query):
        name = self.Connection.getObjectNames()
        return len(query) > 0 and name.isNameValid(QUERY, query) and not name.isNameUsed(QUERY, query)

    def setQuery(self, query, subquery):
        saved = False
        if self._query is not None and self._subquery is not None:
            saved = self._saveQuery(self._query, self._composer)
            saved |= self._saveQuery(self._subquery.First, self._subcomposer)
        if saved:
            print("MergerModel.setQuery() *******************************************")
            self._addressbook.DatabaseDocument.store()
        # FIXME: If we want to be able to use XSingleSelectQueryAnalyzer::getOrderColumns()
        # FIXME: we must initialize in this order: clear the Order and set a Query
        self._subcomposer.setOrder('')
        subcommand = self._queries.getByName(subquery.First).Command
        self._subcomposer.setQuery(subcommand)
        count = self._subcomposer.getOrderColumns().getCount()
        command = self._queries.getByName(query).Command
        self._composer.setQuery(command)
        self._query = query
        self._subquery = subquery
        self._identifiers, self._emails = self._getSubQueryInfos()
        print("MergerModel.setQuery() %s - %s - Table: %s" % (self._identifiers, self._emails, self._subquery.Second))
        return self._identifiers, self._emails
        #command = self._getQueryCommand(subquery)
        # FIXME: com.sun.star.sdb.RowSet.UpdateTableName 
        #self._address.UpdateTableName = command
        #self._subcomposer.setQuery(command)
        #self._setRowSetFilter(self._address, self._subcomposer)
        #self._recipient.Command = subquery
        #command = self._getQueryCommand(query)
        #self._recipient.Command = query
        #self._recipient.UpdateTableName = command
        #self._composer.setQuery(command)
        #self._setRowSetFilter(self._recipient, self._composer)

    def _saveQuery(self, name, composer):
        if self._queries.hasByName(name):
            query = self._queries.getByName(name)
            command = composer.getQuery()
            if query.Command != command:
                print("MergerModel._saveQuery() *******************************************")
                query.Command = command
                return True
        return False

    def _getSubQueryInfos(self):
        emails = []
        identifiers = []
        orders = self._subcomposer.getOrderColumns()
        for index in range(orders.getCount()):
            identifiers.append(orders.getByIndex(index).Name)
        filters = self._subcomposer.getStructuredFilter()
        for filter in filters:
            if len(filter) > 0:
                emails.append(filter[0].Name)
        return identifiers, emails

    def addQuery(self, name, table):
        # TODO: If we want to be able to use XObjectNames::suggestName,
        # TODO: we must first create query then the subquery.
        query = self._createQuery(name)
        self._insertQuery(name, query)
        subquery = self._addSubQuery(name, table)
        self._addQuery(query, subquery)
        self._addressbook.DatabaseDocument.store()
        return subquery

    def _createQuery(self, name):
        # FIXME: If a Query already exist we rewrite it content!!!
        if self._queries.hasByName(name):
            query = self._queries.getByName(name)
        else:
            service = 'com.sun.star.sdb.QueryDefinition'
            query = createService(self._ctx, service)
        return query

    def _insertQuery(self, name, query):
        if not self._queries.hasByName(name):
            self._queries.insertByName(name, query)

    def _addQuery(self, query, subquery):
        table = self._datasource.DataBase.getQuotedQueryName(subquery.First)
        command = getSqlQuery(self._ctx, 'getQueryCommand', (table, table))
        if self._similar:
            self._composer.setQuery(command)
            self._composer.setStructuredFilter(self._getQueryNullFilters())
            command = self._composer.getQuery()
        query.Command = command

    def _addSubQuery(self, name, table):
        arg = self._getQuotedTableName(table)
        subquery = self.Connection.getObjectNames().suggestName(QUERY, name)
        command = getSqlQuery(self._ctx, 'getQueryCommand', (arg, arg))
        query = self._createQuery(subquery)
        if self._subquery is not None and self._similar:
            self._subcomposer.setQuery(command)
            self._subcomposer.setOrder(self._getSubQueryOrder())
            command = self._subcomposer.getQuery()
        query.Command = command
        self._insertQuery(subquery, query)
        return uno.createUnoStruct('com.sun.star.beans.StringPair', subquery, table)

    def removeQuery(self, query, subquery, last):
        self._queries.removeByName(query)
        self._queries.removeByName(subquery.First)
        self._addressbook.DatabaseDocument.store()
        if last:
            self._query = self._subquery = None

    # Query private shared method
    def _getQueryCommand(self, name):
        return self._queries.getByName(name).Command

    def _getQuotedTableName(self, table):
        return self._datasource.DataBase.getQuotedTableName(table)

    # Email methods
    def addEmail(self, subquery, email):
        if email not in self._emails:
            self._emails.append(email)
        self._subcomposer.setStructuredFilter(self._getSubQueryFilters())
        return self._emails

    def removeEmail(self, subquery, email):
        if email in self._emails:
            self._emails.remove(email)
        self._subcomposer.setStructuredFilter(self._getSubQueryFilters())
        #self._setRowSetFilter(self._address, self._subcomposer)
        return self._emails

    def moveEmail(self, subquery, email, position):
        if email in self._emails:
            self._emails.remove(email)
            if 0 <= position <= len(self._emails):
                self._emails.insert(position, email)
        self._subcomposer.setStructuredFilter(self._getSubQueryFilters())
        return self._emails

    # Email private shared methods
    def _getSubQueryFilters(self):
        filters = []
        for email in self._emails:
            filter = getPropertyValue(email, 'IS NOT NULL', 0, NOT_SQLNULL)
            filters.append((filter, ))
        return tuple(filters)

    # Identifier methods
    def addIdentifier(self, query, subquery, table, identifier):
        if identifier not in self._identifiers:
            self._identifiers.append(identifier)
        self._composer.setStructuredFilter(self._getQueryNullFilters())
        self._subcomposer.setOrder(self._getSubQueryOrder())
        return self._identifiers

    def removeIdentifier(self, query, subquery, identifier):
        if identifier in self._identifiers:
            self._identifiers.remove(identifier)
        self._composer.setStructuredFilter(self._getQueryNullFilters())
        self._subcomposer.setOrder(self._getSubQueryOrder())
        return self._identifiers

    def moveIdentifier(self, query, subquery, identifier, position):
        if identifier in self._identifiers:
            self._identifiers.remove(identifier)
            if 0 <= position <= len(self._identifiers):
                self._identifiers.insert(position, identifier)
        self._composer.setStructuredFilter(self._getQueryNullFilters())
        self._subcomposer.setOrder(self._getSubQueryOrder())
        return self._identifiers

    def _getQueryNullFilters(self):
        filters = []
        for identifier in self._identifiers:
            filter = getPropertyValue(identifier, 'IS NULL', 0, SQLNULL)
            filters.append(filter)
        return (tuple(filters), )

    def _getSubQueryOrder(self):
        orders = []
        for identifier in self._identifiers:
            orders.append('"%s"' % identifier)
        order = ', '.join(orders)
        print("MergerModel._getSubQueryOrder() Order: %s" % order)
        return order

    def _getQueryOrder(self):
        orders = []
        for identifier in self._identifiers:
            orders.append('"%s"' % identifier)
        order = ', '.join(orders)
        print("MergerModel._getQueryOrder() Order: %s" % order)
        return order

    # Email and Identifier shared methods
    def isSimilar(self):
        return self._similar

    # WizardPage1 commitPage()
    def commitPage1(self):
        saved1 = self._saveQuery(self._query, self._composer)
        saved2 = self._saveQuery(self._subquery.First, self._subcomposer)
        if saved1 or saved2:
            print("MergerModel.commitPage1() *******************************************")
            self._addressbook.DatabaseDocument.store()
        name = self._addressbook.Name
        if self._isPage2Loaded():
            # This will only be executed if WizardPage2 has already been loaded
            args = None
            if self._name != name:
                self._changed = True
                self._initGrid1RowSet()
                self._initGrid2RowSet()
                args = (self._recipient, )
                # The RowSet self._address will be executed in Page2 with setAddressTable()
            elif saved1 and not saved2:
                self._initGrid1RowSet()
                args = (self._recipient, )
            elif saved2:
                # Update Grid Address only for a change of filters
                self._initGrid1RowSet()
                self._initGrid2RowSet()
                args = (self._address, self._recipient)
            if args is not None:
                Thread(target=self._executeRowSet, args=args).start()
        self._name = name

    def _executeRowSet(self, *rowsets):
        for rowset in rowsets:
            rowset.execute()

    def _isPage2Loaded(self):
        return self._grid1 is not None

# Procedures called by WizardPage2
    def getPageInfos(self):
        tables = self._getTables() if self._similar else self._getSimilarTables()
        message = self._getMailingMessage()
        return tables, self._subquery.Second, self._getTabTitle(1), self._getTabTitle(2), message

    def getChangedPageInfos(self):
        tables = self._getTables() if self._similar else self._getSimilarTables()
        message = self._getMailingMessage()
        return tables, self._subquery.Second, message

    def _getSimilarTables(self):
        columns = self._tables.get(self._subquery.Second, ())
        tables = [table for table in self._tables if self._tables[table] == columns]
        return tuple(tables)

    def initPage2(self, *args):
        # We must prevent a second execution of self._address in setAddressTable()
        self._changed = False
        Thread(target=self._initPage2, args=args).start()

    def commitPage2(self):
        saved1 = self._saveQuery(self._query, self._composer)
        saved2 = self._saveQuery(self._subquery.First, self._subcomposer)
        if saved1 or saved2:
            print("MergerModel.commitPage2() *******************************************")
            self._addressbook.DatabaseDocument.store()

    def setGrid1Data(self, rowset):
        self._grid1.setDataModel(rowset, self._identifiers)

    def setGrid2Data(self, rowset):
        self._grid2.setDataModel(rowset, self._identifiers)
        return self._getFilterCount()

    def _getFilterCount(self):
        filters = self._composer.getStructuredFilter()
        print("MergerModel._getFilterCount() Filters: %s" % (filters, ))
        return len(filters)

    def hasFilters(self):
        filters = self._composer.getStructuredFilter()
        print("MergerModel.hasFilters() Filters: %s" % (filters, ))
        return len(filters) > 0

    def getGrid1SelectedStructuredFilters(self):
        return self._grid1.getSelectedStructuredFilters()

    def getGrid2SelectedStructuredFilters(self):
        return self._grid2.getSelectedStructuredFilters()

    def hasPendingChange(self):
        return self._changed

    def setAddressTable(self, table):
        if self._changed or self._address.Command != table:
            self._changed = False
            self._address.Command = table
            Thread(target=self._executeAddress).start()

    def getAddressCount(self):
        return self._address.RowCount

    def getRecipientCount(self):
        return self._recipient.RowCount

    def addItem(self, *args):
        Thread(target=self._addItem, args=args).start()

    def addAllItem(self, *args):
        Thread(target=self._addAllItem, args=args).start()

    def removeItem(self, *args):
        self._grid2.deselectAllRows()
        Thread(target=self._removeItem, args=args).start()

    def removeAllItem(self, *args):
        self._grid2.deselectAllRows()
        Thread(target=self._removeAllItem, args=args).start()

    def setAddressRecord(self, index):
        row = self._grid1.getUnsortedIndex(index) +1
        self._setDocumentRecord(self._document, self._address, row)

    def setRecipientRecord(self, index):
        row = self._grid2.getUnsortedIndex(index) +1
        self._setDocumentRecord(self._document, self._recipient, row)

    def getMailingMessage(self):
        message = self._getMailingMessage()
        return message

# Private procedures called by WizardPage2
    def _initPage2(self, window1, window2, initPage):
        self._initRowSet()
        self._grid1 = GridManager(self._ctx, self._url, GridModel(self._ctx), window1, 'MergerGrid1', MULTI, None, 8, True)
        self._grid2 = GridManager(self._ctx, self._url, GridModel(self._ctx), window2, 'MergerGrid2', MULTI, None, 8, True)
        initPage(self._grid1, self._grid2, self._address, self._recipient)
        self._address.execute()
        self._recipient.execute()

    def _executeAddress(self):
        self._address.execute()

    def _addItem(self, filters, enableRemoveAll):
        self._updateItem(filters, True)
        enableRemoveAll(False)

    def _addAllItem(self, table):
        tables = self._composer.getTables().getElementNames()
        main = self._datasource.DataBase.getQuotedQueryName(self._subquery.First)
        format = (main, self._getSubQueryTables(tables, table, True))
        command = getSqlQuery(self._ctx, 'getQueryCommand', format)
        self._composer.setQuery(command)
        self._queries.getByName(self._query).Command = self._composer.getQuery()
        self._addressbook.DatabaseDocument.store()
        self._executeRecipient()

    def _removeItem(self, filters, enableAddAll):
        count = self._updateItem(filters, False)
        enableAddAll(count < 2)

    def _removeAllItem(self, table):
        tables = self._composer.getTables().getElementNames()
        main = self._datasource.DataBase.getQuotedQueryName(self._subquery.First)
        format = (main, self._getSubQueryTables(tables, table, True))
        command = getSqlQuery(self._ctx, 'getQueryCommand', format)
        self._composer.setQuery(command)
        if len(tables) <= 2:
            self._composer.setStructuredFilter(self._getQueryNullFilters())
        self._queries.getByName(self._query).Command = self._composer.getQuery()
        self._addressbook.DatabaseDocument.store()
        self._executeRecipient()

    def _updateItem(self, items, add):
        filters = self._getComposerFilter(self._composer)
        for item in items:
            self._updateFilters(filters, item, add)
        self._composer.setStructuredFilter(filters)
        self._queries.getByName(self._query).Command = self._composer.getQuery()
        self._addressbook.DatabaseDocument.store()
        self._executeRecipient()
        return len(filters)

    def _getSubQueryTables(self, tables, table, add):
        if self._subquery.Second != table:
            query = self._datasource.DataBase.getInnerJoinTable(self._subquery.First, self._identifiers, tables, table, add)
        else:
            query = self._datasource.DataBase.getQuotedQueryName(self._subquery.First)
        return query

    def _getComposerFilter(self, composer):
        filters = composer.getStructuredFilter()
        return list(filters)

    def _updateFilters(self, filters, filter, add):
        if add:
            if filter not in filters:
                filters.append(filter)
        elif filter in filters:
            filters.remove(filter)

    def _setDocumentRecord(self, document, rowset, index):
        url = None
        if document.supportsService('com.sun.star.text.TextDocument'):
            url = '.uno:DataSourceBrowser/InsertContent'
        elif document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
            url = '.uno:DataSourceBrowser/InsertColumns'
        if url is not None:
            result = rowset.createResultSet()
            if result is not None:
                descriptor = self._getDataDescriptor(result, index)
                frame = document.CurrentController.Frame
                executeFrameDispatch(self._ctx, frame, url, descriptor)

    def _getDataDescriptor(self, result, row):
        # FIXME: We need to provide ActiveConnection, DataSourceName, Command and CommandType parameters,
        # FIXME: but apparently only Cursor, BookmarkSelection and Selection parameters are used!!!
        properties = {'ActiveConnection': self.Connection,
                      'DataSourceName': self._addressbook.Name,
                      'Command': self._subquery.Second,
                      'CommandType': TABLE,
                      'Cursor': result,
                      'BookmarkSelection': False,
                      'Selection': (row, )}
        descriptor = getPropertyValueSet(properties)
        return descriptor

# Procedures called by WizardPage1 and WizardPage2
    def _initRowSet(self):
        self._initGrid1RowSet()
        self._initGrid2RowSet()

    def _initGrid1RowSet(self):
        self._address.ApplyFilter = False
        self._address.Command = self._subquery.Second
        self._address.Filter = self._subcomposer.getFilter()
        self._address.ApplyFilter = True

    def _initGrid2RowSet(self):
        self._recipient.Command = self._query
        self._recipient.Order = self._getQueryOrder()

# Procedures called by WizardPage3
    def initPage3(self, *args):
        Thread(target=self._initPage3, args=args).start()

    def getUrl(self):
        return self._document.URL

    def getDocumentInfo(self):
        filters = self._grid2.getGridFilters()
        url = getUrlPresentation(self._ctx, self.getUrl())
        return url, self._name, self._query, self._subquery.Second, filters

    def saveDocument(self):
        self._saved = True
        if not self._document.hasLocation():
            frame = self._document.CurrentController.Frame
            url = '.uno:Save'
            listener = DispatchListener(self)
            executeFrameDispatch(self._ctx, frame, url, (), listener)
        return self._saved

    def saveDocumentFinished(self, saved):
        self._saved = saved

    def getDocument(self, url=None):
        if url is None:
            document = self._document
        else:
            document = getDocument(self._ctx, url)
        return document

    def setUrl(self, url):
        pass

    def getRecipients(self):
        recipients = self._getRecipients()
        message = self._getRecipientMessage(len(recipients))
        return recipients, message

    def _getRecipients(self):
        query = self._datasource.DataBase.getQuotedQueryName(self._query)
        columns = self._datasource.DataBase.getRecipientColumns(self._emails)
        order = self._getQueryOrder()
        format = (columns, query, order)
        command = getSqlQuery(self._ctx, 'getRecipientQuery', format)
        print("MergerModel._getRecipients() Query: %s" % command)
        result = self._statement.executeQuery(command)
        recipients = getSequenceFromResult(result)
        result.close()
        return recipients

    def hasMergeMark(self, url):
        return url.endswith('#merge&pdf') or url.endswith('#merge')

    def mergeDocument(self, document, index):
        self._setDocumentRecord(document, self._recipient, index +1)

# Private procedures called by WizardPage3
    def _initPage3(self, handler, initView, initRecipient):
        self._recipient.addRowSetListener(handler)
        initView(self._document)
        recipients, message = self.getRecipients()
        initRecipient(recipients, message)

    def _getDocumentName(self):
        url = None
        location = self._document.getLocation()
        if location:
            url = getUrl(self._ctx, location)
        return None if url is None else url.Name

# Procedures called internally
    def _getPath(self):
        location = self.getUrl()
        if location != '':
            url = getUrl(self._ctx, location)
            path = url.Protocol + url.Path
        else:
            path = getPathSettings(self._ctx).Work
        return path

    # RowSet private methods
    def _getRowSet(self, command):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.CommandType = command
        rowset.FetchSize = g_fetchsize
        return rowset

    def _executeRecipient(self):
        self._recipient.ApplyFilter = False
        self._recipient.Filter = self._composer.getFilter()
        self._recipient.ApplyFilter = True
        self._recipient.execute()

    def _setRowSetFilter(self, rowset, composer):
        rowset.ApplyFilter = False
        rowset.Filter = composer.getFilter()
        rowset.ApplyFilter = True

    def _setRowSetOrder(self, rowset, composer):
        rowset.Order = composer.getOrder()

# MergerModel StringRessoure methods
    def getPageStep(self, pageid):
        resource = self._resources.get('Step') % pageid
        step = self._resolver.resolveString(resource)
        return step

    def getPageTitle(self, pageid):
        resource = self._resources.get('Title') % pageid
        title = self._resolver.resolveString(resource)
        return title

    def _getTabTitle(self, tab):
        resource = self._resources.get('TabTitle') % tab
        return self._resolver.resolveString(resource)

    def getProgressMessage(self, value):
        resource = self._resources.get('Progress') % value
        return self._resolver.resolveString(resource)

    def _getErrorMessage(self, code, format=()):
        resource = self._resources.get('Error') % code
        return self._resolver.resolveString(resource) % format

    def _getIndexLabel(self):
        resource = self._resources.get('Index') % int(self._similar)
        return self._resolver.resolveString(resource)

    def _getRecipientMessage(self, total):
        resource = self._resources.get('Recipient')
        return self._resolver.resolveString(resource) + '%s' % total

    def _getMailingMessage(self):
        resource = self._resources.get('Message')
        return self._resolver.resolveString(resource) + self._query
