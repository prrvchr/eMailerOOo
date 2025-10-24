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

import uno
import unohelper

from com.sun.star.frame.DispatchResultState import FAILURE
from com.sun.star.frame.DispatchResultState import SUCCESS

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.view.SelectionType import MULTI

from com.sun.star.uno import Exception as UnoException

from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.sdb.SQLFilterOperator import SQLNULL
from com.sun.star.sdb.SQLFilterOperator import NOT_SQLNULL

from com.sun.star.sdb.tools.CompositionType import ForDataManipulation

from ...listener import CloseListener
from ...listener import DispatchListener

from ...grid import GridManager

from ...dialog import MailModel

from ...helper import getConnection
from ...helper import getDataBaseContext
from ...helper import mergeDocument
from ...helper import saveDocumentTo

from ...unotool import StatusIndicator
from ...unotool import TaskEvent

from ...unotool import createService
from ...unotool import executeDesktopDispatch
from ...unotool import executeFrameDispatch
from ...unotool import getConfiguration
from ...unotool import getDocument
from ...unotool import getLastNamedParts
from ...unotool import getNamedValueSet
from ...unotool import getPathSettings
from ...unotool import getPropertyValue
from ...unotool import getPropertyValueSet
from ...unotool import getResourceLocation
from ...unotool import getUrl
from ...unotool import getUrlPresentation
from ...unotool import hasInterface
from ...unotool import saveTopWindowPosition
from ...unotool import setProgress

from ...dbtool import getSequenceFromResult

from ...dbqueries import getSqlQuery

from ...configuration import g_identifier
from ...configuration import g_extension
from ...configuration import g_fetchsize
from ...configuration import g_mergerframe

from threading import Thread
from threading import Lock
from time import sleep
import string
import traceback


class MergerModel(MailModel):
    def __init__(self, ctx, document):
        super().__init__(ctx)
        self._listener = CloseListener(self)
        document.addCloseListener(self._listener)
        self._document = document
        self._url = document.getLocation()
        self._path = self._getPath()
        self._addressbook = None
        self._book = None
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
        self._quoted = True
        self._quote = ''
        self._rows = ()
        self._tables = {}
        self._columns = ()
        self._query = None
        self._subquery = None
        self._identifiers = ()
        self._emails = ()
        self._changes = 0
        self._similar = False
        self._temp = False
        self._saved = False
        self._closing = False
        self._dispatchs = {}
        self._paths = getPathSettings(ctx)
        self._lock = Lock()
        self._service = 'com.sun.star.sdb.SingleSelectQueryComposer'
        self._img = getResourceLocation(ctx, g_identifier, 'img')
        self._resources = {'Step':            'MergerPage%s.Step',
                           'Title':           'MergerPage%s.Title',
                           'TabTitle':        'MergerTab%s.Title',
                           'Progress':        'MergerPage1.Label6.Label.%s',
                           'Error':           'MergerPage1.Label8.Error.%s',
                           'Index':           'MergerPage1.Label14.Label.%s',
                           'Query':           'MergerTab2.Label1.Label',
                           'Recipient':       'MailWindow.Label4.Label',
                           'Document':        'MailWindow.Label8.Label',
                           'Property':        'Mail.Document.Property.%s',
                           'PickerTitle':     'Mail.FilePicker.Title',
                           'DialogTitle':     'MessageBox.Confirm.Title',
                           'DialogMessage':   'MessageBox.Confirm.Message',
                           'MsgBoxTitle':     'MessageBox.Error.Title',
                           'MsgBoxMsg1':      'MessageBox.Error.Message.1',
                           'MsgBoxMsg2':      'MessageBox.Error.Message.2',
                           'MsgBoxMsg3':      'MessageBox.Error.Message.3',
                           'MsgBoxMsg4':      'MessageBox.Error.Message.4'}

    @property
    def Connection(self):
        return self._statement.getConnection()

    def _isConnectionNotClosed(self):
        return self._statement is not None

    def _closeConnection(self):
        connection = self._statement.getConnection()
        self._statement.dispose()
        self._statement = None
        if self._grid1 is not None:
            self._grid1.resetDataModel()
        if self._grid2 is not None:
            self._grid2.resetDataModel()
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

# XCloseListener
    def queryClosing(self, document, ownership):
        if self._listener:
            document.removeCloseListener(self._listener)
            self._document = getDocument(self._ctx, self._url, False)
            self._listener = None

# Procedures called by WizardController
    def dispose(self):
        self._disposed = True
        if self._isConnectionNotClosed():
            self._closeConnection()
        if self._listener:
            self._document.removeCloseListener(self._listener)
        if self._grid1 is not None:
            self._grid1.dispose()
        if self._grid2 is not None:
            self._grid2.dispose()
        super().dispose()

    def saveView(self, position):
        config = getConfiguration(self._ctx, g_identifier, True)
        saveTopWindowPosition(config, position, 'MergerPosition')
        config.dispose()
        if self._grid1 is not None:
            self._grid1.saveColumnSettings()
        if self._grid2 is not None:
            self._grid2.saveColumnSettings()

# Procedures called by WizardPage1
    # AddressBook methods
    def getAvailableAddressBooks(self):
        service = 'com.sun.star.sdb.DatabaseContext'
        return createService(self._ctx, service).getElementNames()

    def getDefaultAddressBook(self):
        book = ''
        if self._loadAddressBook():
            book = self._getDocumentAddressBook(book)
        return book

    def _loadAddressBook(self):
        return self._config.getByName('MergerLoadDataSource')

    def _getDocumentAddressBook(self, book):
        service = 'com.sun.star.text.TextDocument'
        if self._document.supportsService(service):
            service = 'com.sun.star.document.Settings'
            settings = self._document.createInstance(service)
            book = settings.CurrentDatabaseDataSource
        return book

    def isAddressBookNotLoaded(self, book):
        return self._book != book

    def setAddressBook(self, *args):
        Thread(target=self._setAddressBook, args=args).start()

    # AddressBook private methods
    def _setAddressBook(self, book, parent, caller):
        sleep(0.2)
        step = 2
        self._setProgress(caller, 5)
        queries = label = message = None
        self._tables = {}
        self._setProgress(caller, 10)
        # FIXME: If changes have been made then save them...
        if self._queries is not None:
            self._saveQueries()
        self._query = self._subquery = None
        # FIXME: If an addressbook has been loaded we need:
        # FIXME: to dispose all components who use the connection and close the connection
        if self._isConnectionNotClosed():
            self._closeConnection()
        self._setProgress(caller, 20)
        self._book = book
        try:
            datasource = self._getDataSource(book)
            self._setProgress(caller, 30)
            connection = getConnection(self._ctx, datasource, parent)
            self._setProgress(caller, 40)
            if not connection:
                msg = self._getErrorMessage(4)
                raise UnoException(msg, self)
            if self._service not in connection.getAvailableServiceNames():
                msg = self._getErrorMessage(5, self._service)
                raise UnoException(msg, self)
            interface = 'com.sun.star.sdb.tools.XConnectionTools'
            if not hasInterface(connection, interface):
                msg = self._getErrorMessage(6, interface)
                raise UnoException(msg, self)
        except UnoException as e:
            message = self._getErrorMessage(0, book, e.Message)
        else:
            self._setProgress(caller, 50)
            # FIXME: We need to keep the datasource name because sometimes datasource.Name returns
            # FIXME: the URL of the odb file rather than the registered name of the datasource
            self._addressbook = datasource
            self._statement = connection.createStatement()
            self._composer = connection.createInstance(self._service)
            self._subcomposer = connection.createInstance(self._service)
            self._setProgress(caller, 60)
            self._tables, self._similar = self._getTablesInfos(connection)
            self._setProgress(caller, 70)
            self._address.ActiveConnection = connection
            self._recipient.ActiveConnection = connection
            self._setProgress(caller, 80)
            self._queries = datasource.getQueryDefinitions()
            # FIXME: SQLite need quoted identifier but don't supports mixed case quoted identifiers.
            #self._quoted = connection.getMetaData().supportsMixedCaseQuotedIdentifiers()
            self._quote = connection.getMetaData().getIdentifierQuoteString()
            self._labeler = connection.getObjectNames()
            self._nominator = connection.createTableName()
            composer = connection.createInstance(self._service)
            subqueries = self._getSubQueries(composer)
            self._setProgress(caller, 90)
            queries = self._getQueries(composer, subqueries)
            composer.dispose()
            label = self._getIndexLabel()
            step = 3
        self._setProgress(caller, 100)
        data = {'call': 'result', 'step': step, 'message': message, 'label': label,
                'tables': self._getTables(), 'queries': queries}
        self._callback.addCallback(caller, getNamedValueSet(data))

    def _setProgress(self, caller, value):
        setProgress(self._callback, caller, value)

    def _saveQueries(self):
        saved = False
        if self._query is not None:
            saved |= self._saveQuery(self._query, self._composer)
        if self._subquery is not None:
            saved |= self._saveQuery(self._subquery.First, self._subcomposer)
        if saved:
            self._saveAddressBook()

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

    def _getDataSource(self, book):
        dbcontext, location = getDataBaseContext(self._ctx, self, book, self._getErrorMessage, 1)
        if self._temp:
            datasource = self._getTempDataSource(dbcontext, book, sf, location)
        else:
            datasource = dbcontext.getByName(book)
        return datasource

    def _getTempDataSource(self, dbcontext, book, sf, location):
        # FIXME: We can undo all changes if the wizard is canceled
        # FIXME: or abort the Wizard while keeping the work already done
        # FIXME: The wizard must be modified to take into account the Cancel button
        url = self._getTempUrl(book)
        if not sf.exists(url):
            sf.copy(location, url)
        datasource = dbcontext.getByName(url)
        return datasource

    def _getTempUrl(self, book):
        return '%s/%s.odb' % (self._paths.Temp, book)

    def _getSubQueries(self, composer):
        queries = {}
        names = self._queries.getElementNames()
        for name in names:
            table = self._getSubQueryTable(composer, names, name)
            if table is not None:
                queries[name] = uno.createUnoStruct('com.sun.star.beans.StringPair', table, '')
        return queries

    def _getSubQueryTable(self, composer, queries, query):
        table = None
        composer.setCommand(query, QUERY)
        for name in composer.getTables().getElementNames():
            if name in queries:
                table = name
                break
        return table

    def _getQueries(self, composer, subqueries):
        queries = {}
        for name, query in subqueries.items():
            if self._queries.hasByName(query.First):
                command = self._queries.getByName(query.First).Command
                table = self._getQueryTable(composer, command)
                if table is not None:
                    query.Second = table
                    queries[name] = query
        return getNamedValueSet(queries)

    def _getQueryTable(self, composer, command):
        table = None
        composer.setQuery(command)
        for name in composer.getTables().getElementNames():
            table = name
            break
        return table

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
                identifiers, emails, table, columns = self._getQuery(query, subquery)
                enabled = False
            else:
                identifiers = emails = columns = ()
                enabled = table and self.isQueryValid(query)
            self._columns = columns
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
        columns = ()
        tables = self._subcomposer.getTables()
        table = self._subcomposer.getTables().getByName(subquery.Second)
        if table.getColumns().hasElements():
            columns = table.getColumns().getElementNames()
        return self._identifiers, self._emails, subquery.Second, columns

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
        self._saveAddressBook()
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
        self._saveAddressBook()
        self._query = self._subquery = None

    # Query private shared method
    def _getQueryCommand(self, name):
        return self._queries.getByName(name).Command

    def _getQuotedTableName(self, table):
        self._nominator.setComposedName(table, ForDataManipulation)
        return self._nominator.getComposedName(ForDataManipulation, self._quoted)

    # Email methods
    def addEmail(self, email):
        with self._lock:
            if email not in self._emails:
                self._emails.append(email)
            self._subcomposer.setStructuredFilter(self._getSubQueryFilters())
            return self._emails

    def removeEmail(self, email, filter):
        with self._lock:
            if email in self._emails:
                self._emails.remove(email)
            self._subcomposer.setStructuredFilter(self._getSubQueryFilters())
            if filter:
                self._composer.setStructuredFilter(self._getQueryNullFilters())
            return self._emails

    def moveEmail(self, email, position):
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
    def addIdentifier(self, identifier, filter):
        with self._lock:
            if identifier not in self._identifiers:
                self._identifiers.append(identifier)
            self._updateQueries(filter)
            return self._identifiers

    def removeIdentifier(self, identifier, filter):
        with self._lock:
            if identifier in self._identifiers:
                self._identifiers.remove(identifier)
            self._updateQueries(filter)
            return self._identifiers

    def moveIdentifier(self, identifier, filter, position):
        with self._lock:
            if identifier in self._identifiers:
                self._identifiers.remove(identifier)
                if 0 <= position <= len(self._identifiers):
                    self._identifiers.insert(position, identifier)
            self._updateQueries(filter)
            return self._identifiers

    def getMessageBoxData(self, query):
        return self._getExceptionMessage(query), self._getDialogTitle()

    def _updateQueries(self, filter):
        if filter:
            self._composer.setStructuredFilter(self._getQueryNullFilters())
        self._subcomposer.setOrder(self._getSubQueryOrder())

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
            name = self._book
            # FIXME: It is necessary to discern the reloading of the list
            # FIXME: of tables of the address book and the grid (or rowset)
            if self._isPage2Loaded():
                # This will only be executed if WizardPage2 has already been loaded
                if self._name != name:
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
                            self._initGrid1RowSet()
                            args.append(self._address)
                    elif changed:
                        self._initGrid1RowSet()
                        args.append(self._address)
                    if self._isGrid2RowSetChanged():
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
    def getPageInfos(self, resolver):
        tables = self._getTables() if self._similar else self._getSimilarTables()
        return tables, self._getTabTitle(resolver, 1), self._getTabTitle(resolver, 2), self._getQueryLabel(resolver)

    def resetPendingChanges(self):
        self._changes = 0

    def hasQueryChanged(self):
        return self._changes & 1 == 1

    def getQueryLabel(self, resolver):
        return self._getQueryLabel(resolver)

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
        self._grid1.enableColumnSelection(False)
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

    def getGrid1(self):
        return self._grid1

    def getGrid2(self):
        return self._grid2

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
        # XXX: If the document has been previously closed,
        # XXX: then no more DocumentRecords will be defined.
        if self._listener is not None:
            row = self._grid1.getUnsortedIndex(index) + 1
            self._setDocumentRecord(self._document, self._address, row)

    def setRecipientRecord(self, index):
        # XXX: If the document has been previously closed,
        # XXX: then no more DocumentRecords will be defined.
        if self._listener is not None:
            row = self._grid2.getUnsortedIndex(index) + 1
            self._setDocumentRecord(self._document, self._recipient, row)

    def addAddressListener(self, listener):
        self._address.addRowSetListener(listener)

    def addRecipientListener(self, listener):
        self._recipient.addRowSetListener(listener)
        self._recipient.execute()

# Private procedures called by WizardPage2
    def _initPage2(self, window1, window2, caller):
        sleep(0.2)
        self._initRowSet()
        self._grid1 = GridManager(self._ctx, self._img, window1, 'MergerGrid1', MULTI, None, 8, True)
        self._grid2 = GridManager(self._ctx, self._img, window2, 'MergerGrid2', MULTI, None, 8, True)
        data = {'address': self._address, 'recipient': self._recipient, 'table': self._subquery.Second}
        self._callback.addCallback(caller, getNamedValueSet(data))

    def _setAddressTable(self, table, enableAddresTable):
        self._address.Command = table
        # FIXME: RowSet.DataSourceName and RowSet.UpdateTableName will be used by the
        # FIXME: GridManager to detect change in GridColumnModel (ie: columns of the table)
        self._address.UpdateTableName = table
        self._address.execute()
        self._grid1.enableColumnSelection(True)
        enableAddresTable(True)

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
            self._saveAddressBook()
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
            self._saveAddressBook()
        self._executeRecipient()

    def _updateItem(self, items, add):
        filters = self._getComposerFilter(self._composer)
        for item in items:
            self._updateFilters(filters, item, add)
        self._composer.setStructuredFilter(filters)
        self._queries.getByName(self._query).Command = self._composer.getQuery()
        self._saveAddressBook()

    def _getSubQueryTables(self, query, table):
        if self._subquery.Second != table:
            table = self._getQuotedTableName(table)
            subquery = self._getQuotedIdentifier(self._subquery.First)
            identifiers = self._getQuotedIdentifiers(self._identifiers)
            query = self._getInnerJoinTable(subquery, identifiers, table)
        return query

    def _getInnerJoinTable(self, subquery, identifiers, table):
         return subquery + self._getInnerJoinPart(identifiers, table)

    def _getInnerJoinPart(self, identifiers, table):
        query = ' INNER JOIN %s ON ' % table
        conditions = []
        for identifier in identifiers:
            conditions.append('%s = %s.%s' % (identifier, table, identifier))
        return query + ' AND '.join(conditions)

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
        result = rowset.createResultSet()
        if result:
            table = self._subquery.Second
            mergeDocument(self._ctx, document, self.Connection, result, self._book, table, index)

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

    def _initGrid2RowSet(self):
        self._recipient.Command = self._query
        self._recipient.Order = self._subcomposer.getOrder()
        # FIXME: RowSet.DataSourceName and RowSet.UpdateTableName will be used by the
        # FIXME: GridManager to detect change in GridColumnModel (ie: columns of the table)
        self._recipient.UpdateTableName = self._query

# Procedures called by WizardPage3
    def isSubjectValid(self, subject):
        if subject != '':
            try:
                for literal, field, format, conversion in string.Formatter().parse(subject):
                    if field is not None and field not in self._columns:
                        return False
                return True
            except ValueError:
                pass
        return False

    def initPage3(self, *args):
        Thread(target=self._initPage3, args=args).start()

    def getUrl(self):
        if self._document:
            url = self._document.getLocation()
        else:
            url = self._url
        return url

    def getDocumentInfo(self):
        url = getUrlPresentation(self._ctx, self.getUrl())
        predicates = self._grid2.getGridPredicates()
        return url, self._name, self._query, self._subquery.Second, predicates, self._emails, self._identifiers

    def saveDocument(self):
        self._saved = True
        if not self._document.hasLocation():
            frame = self._document.CurrentController.Frame
            listener = DispatchListener(self)
            executeFrameDispatch(self._ctx, frame, '.uno:Save', listener)
        return self._saved

# XDispatchResultListener
    def dispatchFinished(self, notification):
        if notification.State == SUCCESS:
            self._saved = notification.Result

    def closeDocument(self, document):
        if self._listener is None:
            document.close(True)

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

    def hasDispatch(self):
        for event in self._dispatchs.values():
            if not event.isSet():
                return True
        return False

    def getTaskEvent(self, filter):
        if filter not in self._dispatchs:
            self._dispatchs[filter] = TaskEvent()
        event = self._dispatchs[filter]
        event.clear()
        return event

    def endDispatch(self, filter):
        if filter in self._dispatchs:
            self._dispatchs[filter].set()

    def cancelDispatch(self):
        for event in self._dispatchs.values():
            event.set()

    def isClosing(self):
        return self._closing

    def setClosing(self):
        self._closing = True

    def _getRecipients(self):
        return self._grid2.getGridData(self._emails, self._quote, '')

    def _viewDocument(self, selection, url, merge, filter, notifier, event):
        # XXX: Breathe
        sleep(0.2)
        if merge:
            connection = self._recipient.ActiveConnection
            table = self._subquery.Second
            result = self._recipient.createResultSet()
        else:
            connection = table = result = None
        kwargs = {'TaskEvent': event, 'Connection': connection, 'ResultSet': result,
                  'DataSource': self._book, 'Table': table, 'Url': url,
                  'Merge': merge, 'Filter': filter, 'Selection': selection}
        executeDesktopDispatch(self._ctx, 'emailer:GetDocument', notifier, **kwargs)

    def _isMerge(self, selection, url, mark):
        return selection and self._hasMergeMark(url, mark)

    def _hasMergeMark(self, url, mark):
        return url is None or mark.startswith('merge')

# Private procedures called by WizardPage3
    def _initPage3(self, listener, caller):
        self._grid2.addGridDataListener(listener)
        recipients, message = self.getRecipients()
        data = {'call': 'init', 'document': self._document,
                'message': message, 'recipients': getNamedValueSet(recipients)}
        self._callback.addCallback(caller, getNamedValueSet(data))

    def _getDocumentName(self):
        url = None
        location = self._document.getLocation()
        if location:
            url = getUrl(self._ctx, location)
        return None if url is None else url.Name

# Procedures called internally
    def _saveAddressBook(self):
        self._addressbook.DatabaseDocument.store()

    def _getPath(self):
        location = self.getUrl()
        if location != '':
            url = getUrl(self._ctx, location)
            path = url.Protocol + url.Path
        else:
            path = self._paths.Work
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
        self._recipient.UpdateTableName = self._recipient.Command
        self._recipient.execute()

    def _setRowSetFilter(self, rowset, composer):
        rowset.ApplyFilter = False
        rowset.Filter = composer.getFilter()
        rowset.ApplyFilter = True

    def _setRowSetOrder(self, rowset, composer):
        rowset.Order = composer.getOrder()

# MergerModel StringRessoure methods
    def getPageStep(self, resolver, pageid):
        resource = self._resources.get('Step') % pageid
        step = resolver.resolveString(resource)
        return step

    def getPageTitle(self, resolver, pageid):
        resource = self._resources.get('Title') % pageid
        title = resolver.resolveString(resource)
        return title

    def _getTabTitle(self, resolver, tab):
        resource = self._resources.get('TabTitle') % tab
        return resolver.resolveString(resource)

    def getProgressMessage(self, resolver, value):
        resource = self._resources.get('Progress') % value
        return resolver.resolveString(resource)

    def _getErrorResource(self, code):
        return self._resources.get('Error') % code

    def _getErrorMessage(self, code, *format):
        resource = self._resources.get('Error') % code
        return self._resolver.resolveString(resource) % format

    def _getIndexLabel(self):
        resource = self._resources.get('Index') % int(self._similar)
        return self._resolver.resolveString(resource)

    def _getRecipientMessage(self, total):
        resource = self._resources.get('Recipient')
        return self._resolver.resolveString(resource) + '%s' % total

    def _getQueryLabel(self, resolver):
        resource = self._resources.get('Query')
        return resolver.resolveString(resource) + self._query

    def _getDialogTitle(self):
        resource = self._resources.get('DialogTitle')
        return self._resolver.resolveString(resource)

    def _getExceptionMessage(self, query):
        resource = self._resources.get('DialogMessage')
        return self._resolver.resolveString(resource) % query

