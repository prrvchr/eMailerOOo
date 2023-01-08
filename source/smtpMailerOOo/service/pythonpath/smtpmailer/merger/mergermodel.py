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

from com.sun.star.sdb.tools.CompositionType import ForDataManipulation
from com.sun.star.sdb.tools.CompositionType import ForTableDefinitions

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
from ..unotool import hasInterface

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
        self._labeler = None
        self._nominator = None
        self._quoted = False
        self._quote = ''
        self._rows = ()
        self._tables = {}
        self._query = None
        self._subquery = None
        self._identifiers = ()
        self._emails = ()
        self._changes = 0
        self._disposed = False
        self._similar = False
        self._temp = False
        self._saved = False
        self._lock = Condition()
        self._service = 'com.sun.star.sdb.SingleSelectQueryComposer'
        self._url = getResourceLocation(ctx, g_identifier, g_extension)
        self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._resources = {'Step': 'MergerPage%s.Step',
                           'Title': 'MergerPage%s.Title',
                           'TabTitle': 'MergerTab%s.Title',
                           'Progress': 'MergerPage1.Label6.Label.%s',
                           'Error': 'MergerPage1.Label8.Label.%s',
                           'Index': 'MergerPage1.Label14.Label.%s',
                           'Query': 'MergerTab2.Label1.Label',
                           'Recipient': 'MailWindow.Label4.Label',
                           'PickerTitle': 'Mail.FilePicker.Title',
                           'Property': 'Mail.Document.Property.%s',
                           'Document': 'MailWindow.Label8.Label.1',
                           'DialogTitle': 'MessageBox.Title',
                           'DialogMessage': 'MessageBox.Message'}

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
            self._query = self._subquery = None
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
                if self._service not in connection.getAvailableServiceNames():
                    msg = self._getErrorMessage(2, self._service)
                    e = self._getUnoException(msg)
                    raise e
                interface = 'com.sun.star.sdb.tools.XConnectionTools'
                if not hasInterface(connection, interface):
                    msg = self._getErrorMessage(3, interface)
                    e = self._getUnoException(msg)
                    raise e
            except UnoException as e:
                format = (addressbook, e.Message)
                message = self._getErrorMessage(0, format)
            else:
                progress(50)
                self._addressbook = datasource
                self._statement = connection.createStatement()
                self._composer = connection.createInstance(self._service)
                self._subcomposer = connection.createInstance(self._service)
                progress(60)
                self._tables, self._similar = self._getTablesInfos(connection)
                progress(70)
                self._address.ActiveConnection = connection
                self._recipient.ActiveConnection = connection
                progress(80)
                self._queries = datasource.getQueryDefinitions()
                self._quoted = connection.getMetaData().supportsMixedCaseQuotedIdentifiers()
                self._quote = connection.getMetaData().getIdentifierQuoteString()
                self._labeler = connection.getObjectNames()
                self._nominator = connection.createTableName()
                composer = connection.createInstance(self._service)
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

    def _saveQuery(self, name, composer):
        if self._queries.hasByName(name):
            query = self._queries.getByName(name)
            command = composer.getQuery()
            if query.Command != command:
                query.Command = command
                return True
        return False

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
        if tables.getCount() > 0:
            table = tables.getElementNames()[0]
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
    def _noFilterChanged(self):
        command = self._queries.getByName(self._query).Command
        return command == self._composer.getQuery()

    def isQueryValid(self, query):
        return len(query) > 0 and self._labeler.isNameValid(QUERY, query) and not self._labeler.isNameUsed(QUERY, query)

    def setQuery(self, *args):
        Thread(target=self._setQuery, args=args).start()

    def _setQuery(self, query, subquery, exist, table, setQuery):
        # FIXME: XComboBox can call the event 'on-textchange' multiple time:
        # FIXME: we have to keep code execution order...
        with self._lock:
            if exist:
                identifiers, emails, table = self._getQuery(query, subquery)
                enabled = False
            else:
                identifiers = emails = ()
                enabled = table and self.isQueryValid(query)
            setQuery(identifiers, emails, exist, table, enabled)

    def _getQuery(self, query, subquery):
        self._saveQueries()
        subcommand = self._queries.getByName(subquery.First).Command
        command = self._queries.getByName(query).Command
        self._subcomposer.setQuery(subcommand)
        self._composer.setQuery(command)
        self._query = query
        self._subquery = subquery
        self._identifiers, self._emails = self._getSubQueryInfos()
        return self._identifiers, self._emails, subquery.Second

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
        composer = None
        query = self._createQuery(name)
        self._insertQuery(name, query)
        if self._similar:
            composer = self.Connection.createInstance(self._service)
        subquery = self._addSubQuery(composer, name, table)
        self._addQuery(composer, query, subquery)
        if self._similar:
            composer.dispose()
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

    def _getQuotedIdentifier(self, identifier):
        return self._quote + identifier + self._quote if self._quoted else identifier

    def _getQuotedIdentifiers(self, identifiers):
        return tuple(map(self._getQuotedIdentifier, identifiers))

    def _addQuery(self, composer, query, subquery):
        table = self._getQuotedIdentifier(subquery.First)
        command = getSqlQuery(self._ctx, 'getQueryCommand', (table, table))
        if self._similar:
            composer.setQuery(command)
            composer.setStructuredFilter(self._getQueryNullFilters())
            command = composer.getQuery()
        query.Command = command

    def _addSubQuery(self, composer, name, table):
        arg = self._getQuotedTableName(table)
        subquery = self._labeler.suggestName(QUERY, name)
        command = getSqlQuery(self._ctx, 'getQueryCommand', (arg, arg))
        query = self._createQuery(subquery)
        if self._similar:
            composer.setQuery(command)
            composer.setOrder(self._getSubQueryOrder())
            command = composer.getQuery()
        query.Command = command
        self._insertQuery(subquery, query)
        return uno.createUnoStruct('com.sun.star.beans.StringPair', subquery, table)

    def removeQuery(self, query, subquery):
        self._queries.removeByName(query)
        self._queries.removeByName(subquery.First)
        self._addressbook.DatabaseDocument.store()
        self._query = self._subquery = None

    # Query private shared method
    def _getQueryCommand(self, name):
        return self._queries.getByName(name).Command

    def _getQuotedTableName(self, table):
        self._nominator.setComposedName(table, ForDataManipulation)
        return self._nominator.getComposedName(ForDataManipulation, self._quoted)

    # Email methods
    def addEmail(self, subquery, email):
        with self._lock:
            if email not in self._emails:
                self._emails.append(email)
            self._subcomposer.setStructuredFilter(self._getSubQueryFilters())
            return self._emails

    def removeEmail(self, subquery, email, count):
        with self._lock:
            if email in self._emails:
                self._emails.remove(email)
            self._subcomposer.setStructuredFilter(self._getSubQueryFilters())
            if count > 1:
                self._composer.setStructuredFilter(self._getQueryNullFilters())
            return self._emails

    def moveEmail(self, subquery, email, position):
        with self._lock:
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
    def addIdentifier(self, query, subquery, table, identifier, count):
        with self._lock:
            if identifier not in self._identifiers:
                self._identifiers.append(identifier)
            self._updateQueries(count)
            return self._identifiers

    def removeIdentifier(self, query, subquery, identifier, count):
        with self._lock:
            if identifier in self._identifiers:
                self._identifiers.remove(identifier)
            self._updateQueries(count)
            return self._identifiers

    def moveIdentifier(self, query, subquery, identifier, position, count):
        with self._lock:
            if identifier in self._identifiers:
                self._identifiers.remove(identifier)
                if 0 <= position <= len(self._identifiers):
                    self._identifiers.insert(position, identifier)
            self._updateQueries(count)
            return self._identifiers

    def getMessageBoxData(self, query):
        return self._getDialogMessage(query), self._getDialogTitle()

    def _updateQueries(self, count):
        if count == 0:
            query = self._getQuotedIdentifier(self._subquery.First)
            format = (query, self._getSubQueryTables(query, self._getInnerTable()))
            command = getSqlQuery(self._ctx, 'getQueryCommand', format)
            self._composer.setQuery(command)
        elif count > 1:
            self._composer.setStructuredFilter(self._getQueryNullFilters())
        self._subcomposer.setOrder(self._getSubQueryOrder())

    def _getInnerTable(self):
        table = self._subquery.Second
        for name in self._composer.getTables().getElementNames():
            if name != self._subquery.First:
                table = name
                break
        print("MergerModel._getInnerTable() SubQuery: %s - Inner Table: %s" % (self._subquery.First, table))
        return table

    def _getQueryNullFilters(self):
        filters = []
        for identifier in self._identifiers:
            filter = getPropertyValue(identifier, 'IS NULL', 0, SQLNULL)
            filters.append(filter)
        return (tuple(filters), )

    def _getSubQueryOrder(self):
        return ', '.join(self._getQuotedIdentifiers(self._identifiers))

    # Email and Identifier shared methods
    def isSimilar(self):
        return self._similar

    # WizardPage1 commitPage()
    def commitPage1(self):
        args = []
        with self._lock:
            self._saveQueries()
            name = self._addressbook.Name
            # FIXME: It is necessary to discern the reloading of the list
            # FIXME: of tables of the address book and the grid (or rowset)
            if self._isPage2Loaded():
                # This will only be executed if WizardPage2 has already been loaded
                if self._name != name:
                    print("MergerModel.commitPage1() DataSource changed")
                    # FIXME: The RowSet self._address will be executed in Page2 by
                    # FIXME: setAddressTable() after selecting the main table
                    self._changes |= 2
                    self._initGrid1RowSet()
                    self._initGrid2RowSet()
                    args.append(self._recipient)
                else:
                    changed = self._isGrid1RowSetChanged()
                    if self._isQueryChanged():
                        # FIXME: If the tables do not all have the same columns (not similar),
                        # FIXME: the list of tables is reloaded each time the mailing list is changed.
                        self._changes |= 1
                        if not self._similar:
                            # FIXME: The RowSet self._address will be executed in Page2 by
                            # FIXME: setAddressTable() after selecting the main table
                            self._changes |= 2
                            self._initGrid1RowSet()
                        elif changed:
                            print("MergerModel.commitPage1() Grid1 changed")
                            self._initGrid1RowSet()
                            args.append(self._address)
                    elif changed:
                        print("MergerModel.commitPage1() Grid1 changed")
                        self._initGrid1RowSet()
                        args.append(self._address)
                    if self._isGrid2RowSetChanged():
                        print("MergerModel.commitPage1() Grid2 changed")
                        self._initGrid2RowSet()
                        args.append(self._recipient)
            self._name = name
        if args:
            Thread(target=self._executeRowSet, args=args).start()

    def _executeRowSet(self, *args):
        for rowset in args:
            rowset.execute()

    def _isPage2Loaded(self):
        return self._grid1 is not None

# Procedures called by WizardPage2
    def getPageInfos(self):
        tables = self._getTables() if self._similar else self._getSimilarTables()
        return tables, self._getTabTitle(1), self._getTabTitle(2), self._getQueryLabel()

    def resetPendingChanges(self):
        self._changes = 0

    def hasQueryChanged(self):
        return self._changes & 1 == 1

    def getQueryLabel(self):
        return self._getQueryLabel()

    def hasTablesChanged(self):
        return self._changes & 2 == 2

    def getQueryTables(self):
        tables = self._getTables() if self._similar else self._getSimilarTables()
        return tables, self._subquery.Second

    def _getSimilarTables(self):
        columns = self._tables.get(self._subquery.Second, ())
        tables = [table for table in self._tables if self._tables[table] == columns]
        return tuple(tables)

    def initPage2(self, *args):
        Thread(target=self._initPage2, args=args).start()

    def setAddressTable(self, *args):
        Thread(target=self._setAddressTable, args=args).start()

    def setGrid1Data(self, rowset):
        self._grid1.deselectAllRows()
        self._grid1.setDataModel(rowset, self._identifiers)

    def setGrid2Data(self, rowset):
        self._grid2.deselectAllRows()
        self._grid2.setDataModel(rowset, self._identifiers)

    def getFilterCount(self):
        filters = self._composer.getStructuredFilter()
        return len(filters)

    def hasFilters(self):
        filters = self._composer.getStructuredFilter()
        return len(filters) > 0

    def getGrid1SelectedStructuredFilters(self):
        return self._grid1.getSelectedStructuredFilters()

    def getGrid2SelectedStructuredFilters(self):
        return self._grid2.getSelectedStructuredFilters()

    def canAdvancePage2(self):
        return self._recipient.RowCount > 0 and self._noFilterChanged()

    def addItem(self, *args):
        Thread(target=self._addItem, args=args).start()

    def addAllItem(self, *args):
        Thread(target=self._addAllItem, args=args).start()

    def removeItem(self, *args):
        self._grid2.deselectAllRows()
        Thread(target=self._removeItem, args=args).start()

    def removeAllItem(self):
        self._grid2.deselectAllRows()
        Thread(target=self._removeAllItem).start()

    def setAddressRecord(self, index):
        row = self._grid1.getUnsortedIndex(index) +1
        self._setDocumentRecord(self._document, self._address, row)

    def setRecipientRecord(self, index):
        row = self._grid2.getUnsortedIndex(index) +1
        self._setDocumentRecord(self._document, self._recipient, row)

# Private procedures called by WizardPage2
    def _initPage2(self, window1, window2, initPage):
        self._initRowSet()
        self._grid1 = GridManager(self._ctx, self._url, GridModel(self._ctx), window1, 'MergerGrid1', MULTI, None, 8, True)
        self._grid2 = GridManager(self._ctx, self._url, GridModel(self._ctx), window2, 'MergerGrid2', MULTI, None, 8, True)
        initPage(self._grid1, self._grid2, self._address, self._recipient, self._subquery.Second)
        self._recipient.execute()

    def _setAddressTable(self, table):
        self._address.Command = table
        self._address.execute()

    def _addItem(self, filters):
        with self._lock:
            self._updateItem(filters, True)
        self._executeRecipient()

    def _addAllItem(self, table):
        with self._lock:
            query = self._getQuotedIdentifier(self._subquery.First)
            format = (query, self._getSubQueryTables(query, table))
            command = getSqlQuery(self._ctx, 'getQueryCommand', format)
            self._composer.setQuery(command)
            self._queries.getByName(self._query).Command = self._composer.getQuery()
            self._addressbook.DatabaseDocument.store()
        self._executeRecipient()

    def _removeItem(self, filters):
        with self._lock:
            self._updateItem(filters, False)
        self._executeRecipient()

    def _removeAllItem(self):
        with self._lock:
            query = self._getQuotedIdentifier(self._subquery.First)
            format = (query, query)
            command = getSqlQuery(self._ctx, 'getQueryCommand', format)
            self._composer.setQuery(command)
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

    def _getSubQueryTables(self, query, table):
        if self._subquery.Second != table:
            table = self._getQuotedTableName(table)
            subquery = self._getQuotedIdentifier(self._subquery.First)
            identifiers = self._getQuotedIdentifiers(self._identifiers)
            query = self._datasource.DataBase.getInnerJoinTable(subquery, identifiers, table)
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

    def _isQueryChanged(self):
        return self._recipient.Command != self._query

    def _isGrid1RowSetChanged(self):
        # Update Grid Address only for a change of Identifiers (order) or Emails (filter)
        return (self._address.Filter != self._subcomposer.getFilter() or
                self._address.Order != self._subcomposer.getOrder())

    def _isGrid2RowSetChanged(self):
        return (self._address.Filter != self._subcomposer.getFilter() or
                self._recipient.Order != self._subcomposer.getOrder() or
                self._recipient.ActiveCommand != self._composer.getQuery())

    def _initGrid1RowSet(self):
        self._address.ApplyFilter = False
        self._address.Filter = self._subcomposer.getFilter()
        self._address.Order = self._subcomposer.getOrder()
        self._address.ApplyFilter = True
        # FIXME: RowSet.DataSourceName and RowSet.UpdateTableName will be used by the
        # FIXME: GridManager to detect change in GridColumnModel (ie: columns of the table)
        self._address.UpdateTableName = self._subquery.Second

    def _initGrid2RowSet(self):
        self._recipient.Command = self._query
        self._recipient.Order = self._subcomposer.getOrder()
        # FIXME: RowSet.DataSourceName and RowSet.UpdateTableName will be used by the
        # FIXME: GridManager to detect change in GridColumnModel (ie: columns of the table)
        self._recipient.UpdateTableName = self._subquery.Second

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
        query = self._getQuotedIdentifier(self._query)
        emails = self._getQuotedIdentifiers(self._emails)
        columns = self._datasource.DataBase.getRecipientColumns(emails)
        order = self._subcomposer.getOrder()
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
        # FIXME: If we want the RowSet listeners to be notified,
        # FIXME: we need to disable then re-enable RowSet.ApplyFilter
        self._recipient.ApplyFilter = False
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

    def _getQueryLabel(self):
        resource = self._resources.get('Query')
        return self._resolver.resolveString(resource) + self._query

    def _getDialogTitle(self):
        resource = self._resources.get('DialogTitle')
        return self._resolver.resolveString(resource)

    def _getDialogMessage(self, query):
        resource = self._resources.get('DialogMessage')
        return self._resolver.resolveString(resource) % query

