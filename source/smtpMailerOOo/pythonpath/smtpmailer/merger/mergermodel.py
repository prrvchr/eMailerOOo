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
from ..unotool import getSimpleFile
from ..unotool import getStringResource
from ..unotool import getUrl
from ..unotool import getUrlPresentation

from ..dbtool import getObjectSequenceFromResult
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
        self._resultset = None
        self._composer = None
        self._subcomposer = None
        self._command = None
        self._subcommand = None
        self._name = None
        self._rows = ()
        self._tables = {}
        self._table = None
        self._query = None
        self._changed = False
        self._disposed = False
        self._similar = False
        self._temp = False
        self._saved = False
        self._lock = Condition()
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
        if self._resultset is not None:
            self._resultset.dispose()
            self._resultset = None
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
            self._query = self._table = None
            self._command = self._subcommand = None
            progress(10)
            # FIXME: If an addressbook has been loaded we need to close the connection
            # FIXME: and dispose all components who use the connection
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
                progress(80)
                self._queries = datasource.getQueryDefinitions()
                progress(90)
                queries = self._getQueries()
                label = self._getIndexLabel()
                step = 3
            progress(100)
            setAddressBook(step, queries, self._getTables(), label, message)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

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

    def _getQueries(self):
        names = self._queries.getElementNames()
        queries = [name for name in names if self._hasSubQuery(names, name)]
        return self._getQueryTables(queries)

    def _hasSubQuery(self, names, name):
        hasquery = False
        subquery = self._getSubQueryName(name)
        if not name.startswith(self._prefix) and subquery in names:
            hasquery = self._checkQueries(name, subquery)
        return hasquery

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

    def _getQueryTables(self, queries):
        querytable = {}
        for query in queries:
            subquery = self._getSubQueryName(query)
            self._composer.setCommand(subquery, QUERY)
            tables = self._composer.getTables()
            if tables.hasElements():
                querytable[query] = tables.getElementNames()[0]
        return querytable

    # AddressBook Table methods
    def setAddressBookTable(self, name):
        columns = ()
        tables = self.Connection.getTables()
        if tables.hasByName(name):
            columns = tables.getByName(name).getColumns().getElementNames()
        return columns

    # Query methods
    def isQueryValid(self, query):
        return len(query) > 0 and self.Connection.getObjectNames().isNameValid(QUERY, query)

    def setQuery(self, query):
        subquery = self._getSubQueryName(query)
        self._address.UpdateTableName = subquery
        self._setComposerCommand(self._subcomposer, subquery)
        self._setRowSet(self._address, self._subcomposer)
        #self._recipient.Command = subquery
        self._recipient.Command = query
        self._recipient.UpdateTableName = query
        self._setComposerCommand(self._composer, query)
        self._setRowSet(self._recipient, self._composer)

    def addQuery(self, query, table):
        name = self._getSubQueryName(query)
        self._addSubQuery(name, table)
        command = getSqlQuery(self._ctx, 'getQueryCommand', name)
        self._addQuery(query, command, name)
        self._addressbook.DatabaseDocument.store()

    def removeQuery(self, query):
        self._queries.removeByName(query)
        self._removeSubQuery(query)
        self._addressbook.DatabaseDocument.store()

    # Query private shared method
    def _getSubQueryName(self, query):
        return self._prefix + query

    def _setComposerCommand(self, composer, name):
        query = self._queries.getByName(name)
        command = query.Command
        print("MergerModel._setComposerCommand() '%s'" % query.UpdateTableName)
        composer.setQuery(query.Command)

    def _addSubQuery(self, name, table):
        command = getSqlQuery(self._ctx, 'getQueryCommand', self._getQuotedTableName(table))
        query = self._createQuery(name, command, table)
        self._addQueries(name, query)

    def _getQuotedTableName(self, table):
        return self._datasource.DataBase.getQuotedTableName(table)

    def _addQuery(self, name, command, subquery):
        query = self._createQuery(name, command, subquery)
        if self._similar:
            filters = self._getQueryFilters()
            self._composer.setQuery(command)
            self._composer.setStructuredFilter(filters)
            query.Command = self._composer.getQuery()
        self._addQueries(name, query)

    def _addQueries(self, name, query):
        if not self._queries.hasByName(name):
            self._queries.insertByName(name, query)

    def _createQuery(self, name, command, table=None):
        # FIXME: If a Query already exist we rewrite it content!!!
        if self._queries.hasByName(name):
            query = self._queries.getByName(name)
        else:
            service = 'com.sun.star.sdb.QueryDefinition'
            query = createService(self._ctx, service)
            if table is not None:
                query.UpdateTableName = table
                print("MergerModel._createQuery() '%s'" % query.UpdateTableName)
        query.Command = command
        return query

    def _getQueryFilters(self, null=True):
        f = []
        identifier = self.getIdentifier()
        if identifier is not None:
            filter = self._getIdentifierFilter(identifier, null)
            f.append(filter)
        filters = tuple(f)
        return (filters, )

    def _removeSubQuery(self, query):
        name = self._getSubQueryName(query)
        self._queries.removeByName(name)

    # Email methods
    def getEmails(self):
        emails = []
        filters = self._subcomposer.getStructuredFilter()
        for filter in filters:
            if len(filter) > 0:
                emails.append(filter[0].Name)
        return emails

    def addEmail(self, query, email):
        name = self._getSubQueryName(query)
        emails = self._addEmail(name, email)
        self._addressbook.DatabaseDocument.store()
        return emails

    def removeEmail(self, query, email):
        name = self._getSubQueryName(query)
        emails = self._removeEmail(name, email)
        self._addressbook.DatabaseDocument.store()
        return emails

    def moveEmail(self, query, email, position):
        name = self._getSubQueryName(query)
        emails = self._moveEmail(name, email, position)
        self._addressbook.DatabaseDocument.store()
        return emails

    # Email private shared methods
    def _addEmail(self, name, email):
        emails = self.getEmails()
        if email not in emails:
            emails.append(email)
        filters = self._getSubQueryFilters(emails)
        self._subcomposer.setStructuredFilter(filters)
        self._setRowSetFilter(self._address, self._subcomposer)
        self._setQueryCommand(name, self._subcomposer)
        return emails

    def _removeEmail(self, name, email):
        emails = self.getEmails()
        if email in emails:
            emails.remove(email)
        filters = self._getSubQueryFilters(emails)
        self._subcomposer.setStructuredFilter(filters)
        self._setRowSetFilter(self._address, self._subcomposer)
        self._setQueryCommand(name, self._subcomposer)
        return emails

    def _moveEmail(self, name, email, position):
        emails = self.getEmails()
        if email in emails:
            emails.remove(email)
            if 0 <= position <= len(emails):
                emails.insert(position, email)
        filters = self._getSubQueryFilters(emails)
        self._subcomposer.setStructuredFilter(filters)
        self._setRowSetFilter(self._address, self._subcomposer)
        self._setQueryCommand(name, self._subcomposer)
        return emails

    def _getSubQueryFilters(self, emails):
        filters = []
        for email in emails:
            filter = getPropertyValue(email, 'IS NOT NULL', 0, NOT_SQLNULL)
            filters.append((filter, ))
        return tuple(filters)

    # Identifier methods
    def getIdentifier(self):
        filters = self._composer.getStructuredFilter()
        return self._getFirstFilterName(filters)

    def addIdentifier(self, query, identifier):
        filter = self._getIdentifierFilter(identifier)
        self._setFilter(query, filter)

    def removeIdentifier(self, query, identifier):
        filter = self._getIdentifierFilter(identifier)
        self._setFilter(query, filter, False)

    def _getIdentifierFilter(self, identifier, null=True):
        if null:
            filter = getPropertyValue(identifier, 'IS NULL', 0, SQLNULL)
        else:
            filter = getPropertyValue(identifier, 'IS NOT NULL', 0, NOT_SQLNULL)
        return filter

    # Index private shared methods
    def _getFirstFilterName(self, filters):
        name = None
        if len(filters) > 0:
            filter = filters[0]
            if len(filter):
                name = filter[0].Name
        return name

    def _setFilter(self, query, filter, add=True):
        if self._similar:
            self._setQueriesFilter(query, filter, add)
        self._setQueryFilter(query, filter, add)
        self._setRowSetFilter(self._recipient, self._composer)
        self._addressbook.DatabaseDocument.store()

    def _setQueriesFilter(self, query, filter, add):
        for name in self._queries.getElementNames():
            if not name.startswith(self._prefix) and name != query:
                self._setQueryFilter(name, filter, add)

    def _setQueryFilter(self, name, filter, add):
        query = self._queries.getByName(name)
        self._composer.setQuery(query.Command)
        filters = self._getQueryFilter(filter, add)
        self._composer.setStructuredFilter(filters)
        query.Command = self._composer.getQuery()

    def _getQueryFilter(self, filter, add):
        filters = list(self._composer.getStructuredFilter())
        if len(filters) > 0:
            f = list(filters[0])
            filters[0] = self._getFirstFilter(f, filter, add)
        elif add:
            filters.append((filter, ))
        return tuple(filters)

    def _getFirstFilter(self, filters, filter, add):
        if add:
            if filter not in filters:
                filters.append(filter)
        else:
            for f in reversed(filters):
                if f.Handle == filter.Handle:
                    filters.remove(f)
        return filters

    # Email, Index and Bookmark shared methods
    def isSimilar(self):
        return self._similar

    # Email private shared methods
    def _setQueryCommand(self, name, composer):
        command = composer.getQuery()
        self._queries.getByName(name).Command = command

    # WizardPage1 commitPage()
    def commitPage1(self, query, table):
        name = self._addressbook.Name
        command = self._composer.getQuery()
        subcommand = self._subcomposer.getQuery()
        if self._isPage2Loaded():
            # This will only be executed if WizardPage2 has already been loaded
            args = None
            changed = False if self._similar else self._table != table
            if self._name != name or changed:
                self._changed = True
                self._resultset = None
                # The RowSet self._address will be executed in Page2 with setAddressTable()
                Thread(target=self._executeResultSet).start()
            elif self._query != query:
                self._changed = True
                args = (self._recipient, )
            elif self._subcommand != subcommand:
                # Update Grid Address only for a change of filters
                self._address.Command = table
                args = (self._address, self._recipient)
            elif self._command != command:
                args = (self._recipient, )
            if args is not None:
                Thread(target=self._executeRowSet, args=args).start()
        self._name = name
        self._query = query
        self._table = table
        self._command = command
        self._subcommand = subcommand

    def _executeRowSet(self, *rowsets):
        for rowset in rowsets:
            rowset.execute()

    def _executeResultSet(self):
        self._recipient.execute()
        rowset = self._getRowSet(TABLE)
        rowset.ActiveConnection = self.Connection
        rowset.Command = self._table
        rowset.execute()
        self._resultset = rowset.createResultSet()

    def _isPage2Loaded(self):
        return self._grid1 is not None

# Procedures called by WizardPage2
    def getPageInfos(self):
        tables = self._getTables() if self._similar else self._getSimilarTables()
        message = self._getMailingMessage()
        return tables, self._table, message

    def _getSimilarTables(self):
        columns = self._tables.get(self._table, ())
        tables = [table for table in self._tables if self._tables[table] == columns]
        return tuple(tables)

    def initPage2(self, *args):
        # We must prevent a second execution of self._address in setAddressTable()
        self._changed = False
        Thread(target=self._initPage2, args=args).start()

    def setGrid1Data(self, rowset):
        identifier = self.getIdentifier()
        self._grid1.setDataModel(rowset, identifier)

    def setGrid2Data(self, rowset):
        identifier = self.getIdentifier()
        self._grid2.setDataModel(rowset, identifier)

    def getGrid1SelectedIdentifiers(self):
        return self._grid1.getSelectedIdentifiers()

    def getGrid2SelectedIdentifiers(self):
        identifiers = self._grid2.getSelectedIdentifiers()
        self._grid2.deselectAllRows()
        return identifiers

    def isChanged(self):
        return self._changed

    def setAddressTable(self, table):
        if self.isChanged() or self._address.Command != table:
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
        Thread(target=self._removeItem, args=args).start()

    def setAddressRecord(self, row):
        identifier= self._grid1.getRowIdentifier(row)
        index = self._grid1.getIdentifierIndex()
        if index != -1:
            self._setDocumentRecord(self._document, self._address, identifier, index)

    def setRecipientRecord(self, row):
        identifier= self._grid2.getRowIdentifier(row)
        index = self._grid2.getIdentifierIndex()
        if index != -1:
            self._setDocumentRecord(self._document, self._recipient, identifier, index)

    def getMailingMessage(self):
        message = self._getMailingMessage()
        return message

# Private procedures called by WizardPage2
    def _initPage2(self, table, possize, parent1, parent2, initPage):
        self._address.Command = table
        self._grid1 = GridManager(self._ctx, GridModel(self._ctx), parent1, possize, 'MergerGrid1', None, 8, True, 'Grid1')
        self._grid2 = GridManager(self._ctx, GridModel(self._ctx), parent2, possize, 'MergerGrid2', None, 8, True, 'Grid2')
        initPage(self._grid1, self._grid2, self._address, self._recipient)
        self._executeAddress()
        self._executeResultSet()

    def _executeAddress(self):
        self._address.execute()

    def _addItem(self, identifiers):
        self._updateItem(identifiers, True)

    def _addAllItem(self, table):
        filters = self._subcomposer.getFilter()
        command = getSqlQuery(self._ctx, 'getQueryCommand', self._getQuotedTableName(table))
        self._subcomposer.setQuery(command)
        self._subcomposer.setFilter(filters)
        subquery = self._getSubQueryName(self._query)
        self._queries.getByName(subquery).Command = command
        filters = self._getQueryFilters(False)
        self._updateItemFilter(filters)

    def _removeItem(self, identifiers):
        self._updateItem(identifiers, False)

    def _updateItem(self, identifiers, add):
        filters = self._getFilters(identifiers, add)
        self._updateItemFilter(filters)

    def _updateItemFilter(self, filters):
        self._composer.setStructuredFilter(filters)
        self._queries.getByName(self._query).Command = self._composer.getQuery()
        self._addressbook.DatabaseDocument.store()
        self._setRowSetFilter(self._recipient, self._composer)
        self._recipient.execute()

    def _getFilters(self, identifiers, add):
        filters = self._getComposerFilter(self._composer)
        identifier = self.getIdentifier()
        dbtype = self._grid2.getIdentifierDataType()
        for value in identifiers:
            self._updateFilters(filters, identifier, value, dbtype, add)
        return tuple(filters)

    def _getComposerFilter(self, composer):
        filters = composer.getStructuredFilter()
        return list(filters)

    def _updateFilters(self, filters, identifier, value, dbtype, add):
        filter = self._getFilter(identifier, value, dbtype)
        if add:
            if filter not in filters:
                filters.append(filter)
        elif filter in filters:
            filters.remove(filter)

    def _getFilter(self, identifier, value, dbtype):
        filter = getPropertyValue(identifier, self._datasource.getFilterValue(value, dbtype), 0, EQUAL)
        filters = (filter, )
        return filters

    def _setDocumentRecord(self, document, rowset, identifier, index):
        url = None
        if document.supportsService('com.sun.star.text.TextDocument'):
            url = '.uno:DataSourceBrowser/InsertContent'
        elif document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
            url = '.uno:DataSourceBrowser/InsertColumns'
        if url is not None:
            result = rowset.createResultSet()
            if result is not None:
                 row = self._getIdentifierRow(result, identifier, index)
                 print("MergerModel._setDocumentRecord() Identifier: %s - Index: %s - Row: %s" % (identifier, index, row))
                 if row != -1:
                    descriptor = self._getDataDescriptor(result, row)
                    frame = document.CurrentController.Frame
                    executeFrameDispatch(self._ctx, frame, url, descriptor)

    def _getIdentifierRow(self, result, identifier, index):
        row = -1
        while result.next():
            value = getResultValue(result, index +1)
            if value == identifier:
                row = result.getRow()
                break
        result.beforeFirst()
        return row

    def _getDataDescriptor(self, result, row):
        # FIXME: We need to provide ActiveConnection, DataSourceName, Command and CommandType parameters,
        # FIXME: but apparently only Cursor, BookmarkSelection and Selection parameters are used!!!
        properties = {'ActiveConnection': self.Connection,
                      'DataSourceName': self._addressbook.Name,
                      'Command': self._table,
                      'CommandType': TABLE,
                      'Cursor': result,
                      'BookmarkSelection': False,
                      'Selection': (row, )}
        descriptor = getPropertyValueSet(properties)
        return descriptor

# Procedures called by WizardPage3
    def initPage3(self, *args):
        Thread(target=self._initPage3, args=args).start()

    def getUrl(self):
        return self._document.URL

    def getDocumentInfo(self, identifiers):
        filters = self._getfiltersFromIdentifiers(identifiers)
        url = getUrlPresentation(self._ctx, self.getUrl())
        return url, self._name, self._query, self._table, filters

    def _getfiltersFromIdentifiers(self, identifiers):
        filters = []
        identifier = self.getIdentifier()
        dbtype = self._grid2.getIdentifierDataType()
        for value in identifiers:
            filters.append(self._datasource.getFilter(identifier, value, dbtype))
        return tuple(filters)

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
        emails = self.getEmails()
        columns = getSqlQuery(self._ctx, 'getRecipientColumns', emails)
        identifier = self.getIdentifier()
        filter = self._composer.getFilter()
        command = self._subcomposer.getQuery()
        format = (columns, identifier, command, filter)
        query = getSqlQuery(self._ctx, 'getRecipientQuery', format)
        print("MergerModel._getRecipients() Query: %s" % query)
        result = self._statement.executeQuery(query)
        recipients = getObjectSequenceFromResult(result)
        return recipients

    def hasMergeMark(self, url):
        return url.endswith('#merge&pdf') or url.endswith('#merge')

    def mergeDocument(self, document, identifier):
        index = self._grid2.getIdentifierIndex()
        if index != -1:
            self._setDocumentRecord(document, self._recipient, identifier, index)

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

    def _setRowSet(self, rowset, composer):
        self._setRowSetFilter(rowset, composer)
        self._setRowSetOrder(rowset, composer)

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

    def getTabTitle(self, tab):
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
