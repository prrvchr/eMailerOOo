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

from com.sun.star.sdbc import SQLException

from com.sun.star.document.MacroExecMode import ALWAYS_EXECUTE_NO_WARN

from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.sdb.SQLFilterOperator import EQUAL
from com.sun.star.sdb.SQLFilterOperator import SQLNULL
from com.sun.star.sdb.SQLFilterOperator import NOT_SQLNULL

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpserver import createService
from smtpserver import getConfiguration
from smtpserver import getDesktop
from smtpserver import getInteractionHandler
from smtpserver import getMessage
from smtpserver import getPathSettings
from smtpserver import getPropertyValue
from smtpserver import getPropertyValueSet
from smtpserver import getSequenceFromResult
from smtpserver import getSqlQuery
from smtpserver import getStringResource
from smtpserver import getUrl
from smtpserver import getValueFromResult
from smtpserver import logMessage

from smtpserver import MailModel

from smtpserver import GridModel
from smtpserver import ColumnModel

from smtpserver import g_identifier
from smtpserver import g_extension
from smtpserver import g_fetchsize

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
        self._address = self._getRowSet(COMMAND, g_fetchsize)
        self._recipient = self._getRowSet(COMMAND, g_fetchsize)
        self._column1 = ColumnModel(ctx)
        self._column2 = ColumnModel(ctx)
        self._width1 = self._getColumnsWidth('MergerGrid1Columns')
        self._width2 = self._getColumnsWidth('MergerGrid2Columns')
        self._maxcolumns = 8
        self._composers = {}
        self._queries = None
        self._addressbook = None
        self._tables = ()
        self._table = None
        self._query = None
        self._statement = None
        self._row = 0
        self._disposed = False
        self._similar = False
        self._filtered = False
        self._changed = False
        self._updated = False
        #self._commands = self._getCommands()
        self._lock = Condition()
        self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._resources = {'Step': 'MergerPage%s.Step',
                           'Title': 'MergerPage%s.Title',
                           'TabTitle': 'MergerTab%s.Title',
                           'TabTip': 'MergerTab%s.Tab.ToolTip',
                           'Email': 'MergerPage1.Label11.Label.%s',
                           'Index': 'MergerPage1.Label12.Label.%s',
                           'Mailing': 'MergerTab2.Label1.Label',
                           'Recipient': 'MailWindow.Label4.Label',
                           'PickerTitle': 'Mail.FilePicker.Title',
                           'Property': 'Mail.Document.Property.%s',
                           'Document': 'MailWindow.Label8.Label.1'}

    @property
    def Connection(self):
        return self._statement.getConnection()

# Procedures called by WizardPage1
    # AddressBook methods
    def getAvailableAddressBooks(self):
        return self._dbcontext.getElementNames()

    def getDocumentAddressBook(self):
        addressbook = ''
        service = 'com.sun.star.text.TextDocument'
        if self._document.supportsService(service):
            service = 'com.sun.star.document.Settings'
            settings = self._document.createInstance(service)
            addressbook = settings.CurrentDatabaseDataSource
        return addressbook

    def getDataSource(self):
        return self.Connection.getParent()

    def setAddressBook(self, *args):
        Thread(target=self._setAddressBook, args=args).start()

    # AddressBook Table methods
    def setAddressBookTable(self, table):
        columns = self._getTableColumns(table)
        name = self._getComposerName(table)
        composer = self._getTableComposer(name, table)
        emails = self._getEmails(composer)
        return columns, emails

    # Query methods
    def validateQuery(self, query, exist):
        valid = False
        if not exist and query != '':
            valid = all((query not in self._tables,
                         query not in self._getComposerNames()))
        return valid

    def setQuery(self, query):
        # TODO: We do not save the grid columns for the first change of self._query
        if self._query is not None:
            self._saveColumnWidth(self._column2, self._width2, self._query)
        self._query = query
        composer = self._getQueryComposer()
        #command = self._getRecipientCommand(composer)
        #self._recipient.Command = command
        #self._initColumn2()
        indexes = self._getIndexes(composer)
        self._setChanged()
        return indexes

    def addQuery(self, table, query):
        composers = self._getComposers()
        if query not in composers:
            #name = self._getComposerName(table)
            command = getSqlQuery(self._ctx, 'getComposerCommand', table)
            composer = self._createComposer(command)
            if self._similar:
                indexes = self._getComposerIndexes(composers)
                self._setIndexes(composer, indexes)
            composers[query] = composer

    def removeQuery(self, query):
        composers = self._getComposers()
        if query in composers:
            del composers[query]

    # Email methods
    def addEmail(self, table, email):
        name = self._getComposerName(table)
        composer = self._getTableComposer(name, table)
        emails = self._getEmails(composer)
        if email not in emails:
            emails.append(email)
        if self._similar:
            self._setComposerEmails(emails)
        else:
            self._setEmails(name, composer, emails)
        self._address.Command = composer.getQuery()
        self._filtered = True
        return emails

    def removeEmail(self, table, email):
        name = self._getComposerName(table)
        composer = self._getTableComposer(name, table)
        emails = self._getEmails(composer)
        if email in emails:
            emails.remove(email)
        if self._similar:
            self._setComposerEmails(emails)
        else:
            self._setEmails(name, composer, emails)
        self._address.Command = composer.getQuery()
        self._filtered = True
        return emails

    def moveEmail(self, table, email, index):
        name = self._getComposerName(table)
        composer = self._getTableComposer(name, table)
        emails = self._getEmails(composer)
        if email in emails:
            emails.remove(email)
            if 0 <= index <= len(emails):
                emails.insert(index, email)
        if self._similar:
            self._setComposerEmails(emails)
        else:
            self._setEmails(name, composer, emails)
        self._address.Command = composer.getQuery()
        self._setFiltered()
        return emails

    # Index methods
    def addIndex(self, index):
        composer = self._getQueryComposer()
        indexes = self._getIndexes(composer)
        if index not in indexes:
            indexes.append(index)
        if self._similar:
            self._setComposerIndexes(indexes)
        else:
            self._setIndexes(composer, indexes)
        #command = self._getRecipientCommand(composer)
        #self._recipient.Command = command
        self._setUpdated()
        return indexes

    def removeIndex(self, index):
        composer = self._getQueryComposer()
        indexes = self._getIndexes(composer)
        if index in indexes:
            indexes.remove(index)
        if self._similar:
            self._setComposerIndexes(indexes)
        else:
            self._setIndexes(composer, indexes)
        #command = self._getRecipientCommand(composer)
        #self._recipient.Command = command
        self._setUpdated()
        return indexes

# Procedures called by WizardPage2
    def getFilteredTables(self):
        tables = self.Connection.getTables().getElementNames()
        if not self._similar:
            tables = self._getFilteredTables(tables)
        return tables

    def getGridModels(self, tab, width, factor, tables=None):
        # TODO: com.sun.star.awt.grid.GridColumnModel must be initialized 
        # TODO: before its assignment at com.sun.star.awt.grid.UnoControlGridModel !!!
        if tab == 1:
            data = GridModel(self._address)
            table = self._getDefaultGridTable(tables)
            widths, titles = self._getGridColumns(self._width1, table)
            column = self._column1.getColumnModel(self._address, widths, titles, width, factor)
        elif tab == 2:
            data = GridModel(self._recipient)
            widths, titles = self._getGridColumns(self._width2, self._query)
            column = self._column2.getColumnModel(self._recipient, widths, titles, width, factor)
        return data, column

    def initRowSet(self, address, recipient, initRecipient):
        self._filtered = self._changed = False
        self._address.addRowSetListener(address)
        self._recipient.addRowSetListener(recipient)
        self.initRecipient(initRecipient)

    def initRecipient(self, *args):
        Thread(target=self._initRecipient, args=args).start()

    def _initRecipient(self, initRecipient):
        composer = self._getQueryComposer()
        command = self._getRecipientCommand(composer)
        self._recipient.Command = command
        self._executeRecipient()
        self._initColumn2()
        columns, orders = self._getRecipientColumnsOrders()
        initRecipient(columns, orders)

    def initRecipientGrid(self, *args):
        Thread(target=self._initRecipientGrid, args=args).start()

    def updateRecipient(self):
        Thread(target=self._updateRecipient).start()

    def _updateRecipient(self):
        composer = self._getQueryComposer()
        command = self._getRecipientCommand(composer)
        self._recipient.Command = command
        self._executeRecipient()

    def _setFiltered(self):
        self._filtered = self._updated = True

    def _setChanged(self):
        self._changed = self._updated = True

    def _setUpdated(self):
        self._updated = True

    def isFiltered(self):
        filtered = self._filtered
        self._filtered = False
        return filtered

    def isChanged(self):
        changed = self._changed
        self._changed = False
        return changed

    def isUpdated(self):
        updated = self._updated
        self._updated = False
        return updated

    def setAddressTable(self, table):
        # TODO: We do not save the grid columns for the first change of self._table
        if self._table is not None:
            self._saveColumnWidth(self._column1, self._width1, self._table)
        self._table = table
        name = self._getComposerName(table)
        composer = self._getTableComposer(name, table)
        command = composer.getQuery()
        self.setAddressCommand(command)
        self._initColumn1()
        columns, orders = self._getComposerColumnsOrders(composer)
        return columns, orders

    def setAddressCommand(self, *args):
        Thread(target=self._setAddressCommand, args=args).start()

    def setAddressColumn(self, *args):
        Thread(target=self._setAddressColumn, args=args).start()

    def setAddressOrder(self, *args):
        Thread(target=self._setAddressOrder, args=args).start()

    def getAddressCount(self):
        return self._address.RowCount

    def setRecipientColumn(self, *args):
        Thread(target=self._setRecipientColumn, args=args).start()

    def setRecipientOrder(self, *args):
        Thread(target=self._setRecipientOrder, args=args).start()

    def getRecipientCount(self):
        return self._recipient.RowCount

    def addItem(self, *args):
        Thread(target=self._addItem, args=args).start()

    def removeItem(self, *args):
        Thread(target=self._removeItem, args=args).start()

    def setDocumentRecord(self, row):
        if row != self._row:
            self._row = row
            Thread(target=self._setDocumentRecord).start()

    def getMailingMessage(self):
        composer = self._getQueryComposer()
        table = composer.getTables().getByIndex(0).Name
        message = self._getMailingMessage(table)
        return message

# Procedures called by WizardPage3
    def getUrl(self):
        return self._document.URL

    def getDocument(self, url=None):
        if url is None:
            document = self._document
        else:
            properties = {'Hidden': True, 'MacroExecutionMode': ALWAYS_EXECUTE_NO_WARN}
            descriptor = getPropertyValueSet(properties)
            document = getDesktop(self._ctx).loadComponentFromURL(url, '_blank', 0, descriptor)
        return document

    def setUrl(self, url):
        pass

    def initView(self, *args):
        Thread(target=self._initView, args=args).start()

    def getRecipients(self):
        composer = self._getQueryComposer()
        table = composer.getTables().getByIndex(0).Name
        emails, subquery = self._getComposerInfos(table)
        columns = getSqlQuery(self._ctx, 'getComposerColumns', emails)
        filter = composer.getFilter()
        format = (columns, subquery, filter)
        query = getSqlQuery(self._ctx, 'getComposerQuery', format)
        print("MergerModel.getRecipients() %s" % query)
        result = self._statement.executeQuery(query)
        recipients = getSequenceFromResult(result)
        print("MergerModel.getRecipients() %s" % (recipients, ))
        total = len(recipients)
        message = self._getRecipientMessage(total)
        return recipients, message

    def _getComposerInfos(self, table):
        name = self._getComposerName(table)
        composer = self._getTableComposer(name, table)
        emails = self._getEmails(composer)
        subquery =  self._getComposerSubQuery(table, composer)
        return emails, subquery

# Private procedures called by WizardPage2
    def _initRecipientGrid(self, initRecipient):
        self._initColumn2()
        columns, orders = self._getRecipientColumnsOrders()
        initRecipient(columns, orders)

    def _getRecipientColumnsOrders(self):
        composer = self._getQueryComposer()
        #mri = createService(self._ctx, 'mytools.Mri')
        #mri.inspect(composer)
        columns, orders = self._getComposerColumnsOrders(composer)
        return columns, orders

    def _getColumnsWidth(self, name):
        config = self._configuration.getByName(name)
        width = json.loads(config, object_pairs_hook=OrderedDict)
        return width

    def _getComposerColumnsOrders(self, composer):
        columns = composer.getColumns().getElementNames()
        orders = composer.getOrderColumns().createEnumeration()
        return columns, orders

    def _getGridColumns(self, config, name):
        names = config.get(self._addressbook, {})
        widths = names.get(name, {})
        if widths:
            titles = self._getColumnTitles(widths)
        else:
            columns = self._getDefaultColumns(name)
            titles = self._getColumnTitles(columns)
        return widths, titles

    def _getColumnTitles(self, columns):
        titles = OrderedDict()
        for column in columns:
            titles[column] = column
        return titles

    def _getDefaultColumns(self, name):
        composers = self._getComposers()
        if name in self._tables:
            name = self._getComposerName(name)
        composer = composers[name]
        columns = composer.getColumns().getElementNames()
        return columns[:self._maxcolumns]

    def _setAddressColumn(self, columns, reset):
        if reset:
            columns = self._getDefaultColumns(self._table)
        titles = self._getColumnTitles(columns)
        self._column1.setColumnModel(self._address, titles, reset)

    def _setAddressOrder(self, orders, ascending):
        name = self._getComposerName(self._table)
        composer = self._getTableComposer(name, self._table)
        self._setComposerOrder(composer, orders, ascending)
        command = composer.getQuery()
        self._setQueryCommand(name, command)
        self._setAddressCommand(command)

    def _setRecipientColumn(self, columns, reset):
        if reset:
            columns = self._getDefaultColumns(self._query)
        titles = self._getColumnTitles(columns)
        self._column2.setColumnModel(self._recipient, titles, reset)

    def _setRecipientOrder(self, orders, ascending):
        composer = self._getQueryComposer()
        self._setComposerOrder(composer, orders, ascending)
        command = self._getRecipientCommand(composer)
        self._setRecipientCommand(command)

    def _setComposerOrder(self, composer, orders, ascending):
        olds, news = self._getComposerOrder(composer, orders)
        composer.Order = ''
        for order in olds:
            composer.appendOrderByColumn(order, order.IsAscending)
        columns = composer.getColumns()
        for order in news:
            column = columns.getByName(order)
            composer.appendOrderByColumn(column, ascending)

    def _getComposerOrder(self, composer, news):
        olds = []
        enumeration = composer.getOrderColumns().createEnumeration()
        while enumeration.hasMoreElements():
            column = enumeration.nextElement()
            if column.Name in news:
                olds.append(column)
                news.remove(column.Name)
        return olds, news

    def _addItem(self, rows):
        composer = self._getQueryComposer()
        filters = self._getRowSetFilters(self._address, composer, rows, True)
        composer.setStructuredFilter(filters)
        command = self._getRecipientCommand(composer)
        self._setRecipientCommand(command)

    def _removeItem(self, rows):
        composer = self._getQueryComposer()
        filters = self._getRowSetFilters(self._recipient, composer, rows, False)
        composer.setStructuredFilter(filters)
        command = self._getRecipientCommand(composer)
        self._setRecipientCommand(command)

    def _getRowSetFilters(self, rowset, composer, rows, add):
        indexes = self._getIndexes(composer)
        filters = self._getComposerFilters(composer)
        for row in rows:
            rowset.absolute(row +1)
            self._updateFilters(rowset, indexes, filters, add)
        return self._getStructuredFilters(filters)

    def _getComposerFilters(self, composer):
        filters = []
        for filter in composer.getStructuredFilter():
            filters.append(self._getComposerFilter(filter))
        return filters

    def _getComposerFilter(self, properties):
        filters = []
        for property in properties:
            value = property.Value
            # TODO: Filter can be quoted if its type is String!!!
            if isinstance(value, string_types):
                value = value.strip("'")
            filter = (property.Name, value)
            filters.append(filter)
        return tuple(filters)

    def _updateFilters(self, rowset, indexes, filters, add):
        filter = self._getRowSetFilter(rowset, indexes)
        if add:
            if filter not in filters:
                filters.append(filter)
        elif filter in filters:
            filters.remove(filter)

    def _getRowSetFilter(self, rowset, indexes):
        filters = []
        for index in indexes:
            i = rowset.findColumn(index)
            value = getValueFromResult(rowset, i)
            filter = (index, value)
            filters.append(filter)
        return tuple(filters)

    def _getStructuredFilters(self, filters):
        structured = []
        for filter in filters:
            properties = []
            for name, value in filter:
                operator = EQUAL if value != 'IS NULL' else SQLNULL
                property = getPropertyValue(name, value, None, operator)
                properties.append(property)
            structured.append(tuple(properties))
        return tuple(structured)

    def _setDocumentRecord(self):
        try:
            dispatch = None
            frame = self._document.getCurrentController().Frame
            flag = uno.getConstantByName('com.sun.star.frame.FrameSearchFlag.SELF')
            if self._document.supportsService('com.sun.star.text.TextDocument'):
                url = getUrl(self._ctx, '.uno:DataSourceBrowser/InsertContent')
                dispatch = frame.queryDispatch(url, '_self', flag)
            elif self._document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
                url = getUrl(self._ctx, '.uno:DataSourceBrowser/InsertColumns')
                dispatch = frame.queryDispatch(url, '_self', flag)
            if dispatch is not None:
                descriptor = self._getDataDescriptor()
                dispatch.dispatch(url, descriptor)
        except Exception as e:
            print("MergerModel._setDocumentRecord() ERROR: %s" % traceback.print_exc())

    def _getDataDescriptor(self):
        properties = {'DataSourceName': self._recipient.DataSourceName,
                      'ActiveConnection': self._recipient.ActiveConnection,
                      'Command': self._recipient.Command,
                      'CommandType': self._recipient.CommandType,
                      'Cursor': self._recipient,
                      'Selection': (self._row, ),
                      'BookmarkSelection': False}
        descriptor = getPropertyValueSet(properties)
        return descriptor

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

    def getTabTip(self, tab):
        resource = self._resources.get('TabTip') % tab
        return self._resolver.resolveString(resource)

    def _getEmailLabel(self):
        resource = self._resources.get('Email') % int(self._similar)
        return self._resolver.resolveString(resource)

    def _getIndexLabel(self):
        resource = self._resources.get('Index') % int(self._similar)
        return self._resolver.resolveString(resource)

    def _getRecipientMessage(self, total):
        resource = self._resources.get('Recipient')
        return self._resolver.resolveString(resource) % total

    def _getMailingMessage(self, table):
        resource = self._resources.get('Mailing')
        return self._resolver.resolveString(resource) % (self._query, table)

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
    def _getRowSet(self, command, fetchsize):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.CommandType = command
        #rowset.FetchSize = fetchsize
        print("MergerModel._getRowSet() FetchSize = %s" % rowset.FetchSize)
        return rowset

    def _setAddressCommand(self, command):
        print("MergerModel._setAddressCommand() %s" % command)
        self._address.Command = command
        self._address.execute()

    def _setRecipientCommand(self, command):
        print("MergerModel._setRecipientCommand() %s" % command)
        self._recipient.Command = command
        self._executeRecipient()

    def _getRecipientCommand(self, composer):
        table = composer.getTables().getByIndex(0).Name
        query = '"%s"' % table
        subquery = self._getSubQueryCommand(table)
        command = composer.getQuery().replace(query, subquery)
        print("MergerModel._getRecipientCommand() %s" % command)
        return command

    def _getSubQueryCommand(self, table):
        name = self._getComposerName(table)
        composer = self._getTableComposer(name, table)
        command = self._getComposerSubQuery(table, composer)
        return command

    def _getComposerSubQuery(self, table, composer):
        filter = composer.getFilter()
        format = (table, filter)
        command = getSqlQuery(self._ctx, 'getComposerSubQuery', format)
        return command

    def _executeRecipient(self):
        print("MergerModel._executeRecipient() %s" % self._recipient.Command)
        self._recipient.execute()

    # AddressBook private methods
    def _setAddressBook(self, addressbook, progress, setAddressBook):
        progress(10)
        step = 2
        queries = label1 = label2 = msg = None
        # TODO: We do not save the grid columns for the first load of a new self._addressbook
        if self._addressbook is not None:
            if self._table is not None:
                self._saveColumnWidth(self._column1, self._width1, self._table)
            if self._query is not None:
                self._saveColumnWidth(self._column2, self._width2, self._query)
        self._query = self._table = None
        self._filtered = self._changed = self._updated = False
        sleep(0.2)
        progress(20)
        database = self._getDatabase(addressbook)
        try:
            if database.IsPasswordRequired:
                handler = getInteractionHandler(self._ctx)
                connection = database.connectWithCompletion(handler)
            else:
                connection = database.getConnection('', '')
        except SQLException as e:
            msg = e.Message
        else:
            progress(30)
            service = 'com.sun.star.sdb.SingleSelectQueryComposer'
            if service not in connection.getAvailableServiceNames():
                msg = "MergerModel._setAddressBook() service: %s not available..." % service
            else:
                progress(40)
                self._addressbook = addressbook
                self._statement = connection.createStatement()
                progress(50)
                self._setTablesInfos()
                progress(60)
                self._address.ActiveConnection = connection
                #mri = createService(self._ctx, 'mytools.Mri')
                #mri.inspect(self._recipient)
                self._recipient.ActiveConnection = connection
                #self._recipient.DataSourceName = addressbook
                progress(70)
                self._queries = database.getQueryDefinitions()
                if addressbook in self._composers:
                    composers = self._composers[addressbook]
                else:
                    composers = self._getQueriesComposer()
                    self._composers[addressbook] = composers
                progress(80)
                queries = self._getQueryNames(composers)
                progress(90)
                label1 = self._getEmailLabel()
                label2 = self._getIndexLabel()
                step = 3
        progress(100)
        setAddressBook(step, queries, self._tables, label1, label2, msg)

    def _getDatabase(self, addressbook):
        database = None
        if self._dbcontext.hasByName(addressbook):
            database = self._dbcontext.getByName(addressbook)
        return database

    def _setTablesInfos(self):
        self._similar = True
        tables = self.Connection.getTables()
        self._tables = tables.getElementNames()
        if len(self._tables) > 0:
            columns = tables.getByIndex(0).getColumns().getElementNames()
            for index in range(1, tables.getCount()):
                table = tables.getByIndex(index)
                if columns != table.getColumns().getElementNames():
                    self._similar = False
                    break

    # Composer private methods
    def _getComposers(self):
        return self._composers[self._addressbook]

    def _getTableComposer(self, name, table):
        composers = self._getComposers()
        if name in composers:
            composer = composers[name]
        else:
            command = self._getQueryCommand(name, table)
            composer = self._createComposer(command)
            composers[name] = composer
        return composer

    def _getComposerNames(self):
        names = []
        for table in self._tables:
            name = self._getComposerName(table)
            names.append(name)
        return names

    def _getComposerName(self, table):
        name = '%s%s' % (self._addressbook, table)
        return name

    def _createComposer(self, command):
        service = 'com.sun.star.sdb.SingleSelectQueryComposer'
        composer = self.Connection.createInstance(service)
        composer.setQuery(command)
        return composer

    def _getQueryComposer(self):
        composers = self._getComposers()
        composer = composers[self._query]
        return composer

    def _setComposerEmails(self, emails):
        filter = self._getEmailFilters(emails)
        for table in self._tables:
            name = self._getComposerName(table)
            composer = self._getTableComposer(name, table)
            composer.setStructuredFilter(filter)
            command = composer.getQuery()
            self._setQueryCommand(name, command)

    def _getComposerIndexes(self, composers):
        indexes = []
        queries = self._getQueryNames(composers)
        for query in queries:
            composer = composers[query]
            indexes = self._getIndexes(composer)
            break
        return indexes

    def _setComposerIndexes(self, indexes):
        filter = self._getIndexFilters(indexes)
        composers = self._getComposers()
        queries = self._getQueryNames(composers)
        for query in queries:
            composer = composers[query]
            composer.setStructuredFilter(filter)

    # Queries private methods
    def _getQueriesComposer(self):
        composers = {}
        for name in self._queries.getElementNames():
            query = self._queries.getByName(name)
            composer = self._createComposer(query.Command)
            composers[name] = composer
        return composers

    def _setQueryCommand(self, name, command):
        print("MergerModel._setQueryCommand() 1 %s - %s" % (name, command))
        if self._queries.hasByName(name):
            query = self._queries.getByName(name)
        else:
            service = 'com.sun.star.sdb.QueryDefinition'
            query = createService(self._ctx, service)
            self._queries.insertByName(name, query)
            print("MergerModel._setQueryCommand() 2 %s - %s" % (name, command))
        query.Command = command

    def _getQueryCommand(self, name, table):
        print("MergerModel._getQueryCommand() 1 %s - %s" % (name, table))
        if self._queries.hasByName(name):
            query = self._queries.getByName(name)
            command = query.Command
        else:
            #service = 'com.sun.star.sdb.QueryDefinition'
            #query = createService(self._ctx, service)
            command = getSqlQuery(self._ctx, 'getComposerCommand', table)
            #self._queries.insertByName(name, query)
            #self.Connection.getParent().DatabaseDocument.store()
        print("MergerModel._getQueryCommand() 2 %s - %s - %s" % (name, table, command))
        return command

    def _getQueryNames(self, composers):
        names = self._getComposerNames()
        queries = [name for name in composers if name not in names]
        return tuple(queries)

    def saveQueries1(self):
        self.Connection.getParent().DatabaseDocument.store()

    # Table private methods
    def _getTableColumns(self, table):
        table = self.Connection.getTables().getByName(table)
        columns = table.getColumns().getElementNames()
        return columns

    def _getFilteredTables(self, tables):
        names = []
        composers = self._getComposers()
        for table in tables:
            name = self._getComposerName(table)
            if name in composers:
                composer = composers[name]
                filter = composer.getFilter()
                print("MergerModel._getFilteredTables() '%s'" % filter)
                if filter != '':
                    names.append(table)
        return tuple(names)

    # Email private methods
    def _getEmails(self, composer):
        emails = []
        filters = composer.getStructuredFilter()
        for filter in filters:
            if len(filter) > 0:
                emails.append(filter[0].Name)
        return emails

    def _setEmails(self, name, composer, emails):
        filter = self._getEmailFilters(emails)
        composer.setStructuredFilter(filter)
        command = composer.getQuery()
        self._setQueryCommand(name, command)

    def _getEmailFilters(self, emails):
        filters = []
        for email in emails:
            filter = getPropertyValue(email, 'IS NOT NULL', None, NOT_SQLNULL)
            filters.append((filter, ))
        return tuple(filters)

    # Index private methods
    def _getIndexes(self, composer):
        indexes = []
        filters = composer.getStructuredFilter()
        if len(filters) > 0:
            for filter in filters[0]:
                indexes.append(filter.Name)
        return indexes

    def _setIndexes(self, composer, indexes):
        filter = self._getIndexFilters(indexes)
        composer.setStructuredFilter(filter)

    def _getIndexFilters(self, indexes):
        filters = []
        for index in indexes:
            filter = getPropertyValue(index, 'IS NULL', None, SQLNULL)
            filters.append(filter)
        return (tuple(filters), )

    # Grid Columns private methods
    def _getDefaultGridTable(self, tables):
        return tables[0]

    def _initColumn1(self):
        # TODO: We do not reset the grid columns if ColumnModel is not yet assigned
        if self._column1.isInitialized():
            widths, titles = self._getGridColumns(self._width1, self._table)
            self._column1.initColumnModel(self._address, widths, titles)

    def _initColumn2(self):
        # TODO: We do not reset the grid columns if ColumnModel is not yet assigned
        if self._column2.isInitialized():
            widths, titles = self._getGridColumns(self._width2, self._query)
            self._column2.initColumnModel(self._recipient, widths, titles)

    def _saveColumnWidth(self, column, widths, name):
        # TODO: We do not save the grid columns if ColumnModel is not yet assigned
        if column.isInitialized():
            columns = column.getColumnWidth()
            if self._addressbook not in widths:
                widths[self._addressbook] = {}
            addressbook = widths[self._addressbook]
            addressbook[name] = columns
            print("MergerModel._saveColumnWidth() %s - %s" % (name, widths))

# Private procedures called by WizardPage3
    def _initView(self, handler, initView, initRecipient):
        self._address.addRowSetListener(handler)
        self._recipient.addRowSetListener(handler)
        initView(self._document)
        recipients, message = self.getRecipients()
        initRecipient(recipients, message)



    def _getQuery1(self, queries, name, create):
        print("MergerModel._getQuery() '%s'" % (name, ))
        if queries.hasByName(name):
            query = queries.getByName(name)
        else:
            service = 'com.sun.star.sdb.QueryDefinition'
            query = createService(self._ctx, service)
            if create:
                queries.insertByName(name, query)
            else:
                table = self.Connection.getTables().getByIndex(0)
                column = table.getColumns().getByIndex(0).Name
                format = {'Table': table.Name, 'Column': column}
                query.Command = self._getQueryCommand(format)
                #query.UpdateTableName = table
        return query

    def _getDocumentName(self):
        url = None
        location = self._document.getLocation()
        if location:
            url = getUrl(self._ctx, location)
        return None if url is None else url.Name
