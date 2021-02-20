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

from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService
from unolib import getDesktop
from unolib import getStringResource
from unolib import getPropertyValueSet

from smtpserver.grid import ColumnModel

from smtpserver import g_identifier
from smtpserver import g_extension
from smtpserver import g_fetchsize

from smtpserver import logMessage
from smtpserver import getMessage

from threading import Thread
from time import sleep
import traceback


class MergerModel(unohelper.Base):
    def __init__(self, ctx):
        self._ctx = ctx
        self._stringResource = getStringResource(ctx, g_identifier, g_extension)
        self._doc = getDesktop(ctx).CurrentComponent
        service = 'com.sun.star.sdb.DatabaseContext'
        self._dbcontext = createService(ctx, service)
        self._address = self._getRowSet()
        self._recipient = self._getRowSet()
        self._column1 = ColumnModel(ctx)
        self._column2 = ColumnModel(ctx)
        self._statement = None
        self._composer = None

    @property
    def Connection(self):
        return self._statement.getConnection()

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

# Procedures called by WizardPage1
    def getAvailableDataSources(self):
        return self._dbcontext.getElementNames()

    def getDocumentDataSource(self):
        datasource = ''
        if self._doc.supportsService('com.sun.star.text.TextDocument'):
            service = 'com.sun.star.document.Settings'
            datasource = self._doc.createInstance(service).CurrentDatabaseDataSource
        return datasource

    def getDataSource(self):
        return self.Connection.getParent()

    def setDataSource(self, *args):
        Thread(target=self._setDataSource, args=args).start()

    def getTableColumns(self, table):
        columns = self.Connection.getTables().getByName(table).getColumns().getElementNames()
        return columns

    def getQuery(self, name, create=True):
        print("MergerModel._getQuery() '%s'" % (name, ))
        queries = self.getDataSource().getQueryDefinitions()
        query = self._getQuery(queries, name, create)
        return query

# Private procedures called by WizardPage1
    def _setDataSource(self, datasource, progress, callback):
        progress(10)
        sleep(0.2)
        step = 2
        queries, tables, email, index, msg = None, None, None, None, None
        database = self._getDatabase(datasource)
        progress(20)
        try:
            if database.IsPasswordRequired:
                handler = createService(self._ctx, 'com.sun.star.task.InteractionHandler')
                connection = database.connectWithCompletion(handler)
            else:
                connection = database.getConnection('', '')
        except SQLException as e:
            composer = None
            msg = e.Message
        else:
            progress(30)
            self._statement = connection.createStatement()
            document, form = self._getForm(False)
            progress(40)
            email = self._getDocumentList(document, 'EmailColumns')
            index = self._getDocumentList(document, 'IndexColumns')
            progress(50)
            if form is not None:
                form.close()
            progress(60)
            tbls = self.Connection.getTables()
            #mri = createService(self._ctx, 'mytools.Mri')
            #mri.inspect(tbls)

            self._composer, queries = self._getQueryComposer(progress, g_extension)
            tables = tbls.getElementNames()
            #table = composer.getTables().getByIndex(0)
            step = 3
        progress(100)
        callback(step, queries, tables, email, index, msg)
        #sleep(0.2)
        #if query is not None:
            #self._orders = getOrders(query.getOrder())
            #filter = query.getFilter()
            #self._initRowSet(datasource, tbls.getByIndex(0), filter)

    def _getDatabase(self, datasource):
        database = None
        if self._dbcontext.hasByName(datasource):
            database = self._dbcontext.getByName(datasource)
        return database

    def _getForm(self, create, name=g_extension):
        doc, form = None, None
        forms = self.Connection.getParent().DatabaseDocument.getFormDocuments()
        if forms.hasByName(name):
            form = forms.getByName(name)
        elif create:
            form = self._createForm(forms, name)
        if form is not None:
            args = getPropertyValueSet({'ActiveConnection': self.Connection,
                                        'OpenMode': 'openDesign',
                                        'Hidden': True})
            doc = forms.loadComponentFromURL(name, '', 0, args)
        return doc, form

    def _createForm(self, forms, name):
        service = 'com.sun.star.sdb.DocumentDefinition'
        args = getPropertyValueSet({'Name': name, 'ActiveConnection': self.Connection})
        form = forms.createInstanceWithArguments(service, args)
        forms.insertByName(name, form)
        form = forms.getByName(name)
        return form

    def _getDocumentList(self, document, property):
        items = ()
        value = self._getDocumentValue(document, property)
        if value is not None:
            items = tuple(value.split(','))
        return items

    def _getDocumentValue(self, document, property, default=None):
        value = default
        print("MergerModel._getDocumentValue() %s" % (property, ))
        if document is not None:
            properties = document.DocumentProperties.UserDefinedProperties
            if properties.PropertySetInfo.hasPropertyByName(property):
                print("MergerModel._getDocumentValue() getProperty")
                value = properties.getPropertyValue(property)
            elif default is not None:
                self._setDocumentValue(document, property, default)
        return value

    def _setDocumentValue(self, document, property, value):
        print("MergerModel._setDocumentValue() %s - %s" % (property, value))
        properties = document.DocumentProperties.UserDefinedProperties
        if properties.PropertySetInfo.hasPropertyByName(property):
            print("MergerModel._setDocumentValue() setProperty")
            properties.setPropertyValue(property, value)
        else:
            print("MergerModel._setDocumentValue() addProperty")
            properties.addProperty(property,
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.MAYBEVOID') +
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.BOUND') +
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.REMOVABLE') +
            uno.getConstantByName('com.sun.star.beans.PropertyAttribute.MAYBEDEFAULT'),
            value)

    def _getQueryComposer(self, progress, name):
        service = 'com.sun.star.sdb.SingleSelectQueryComposer'
        composer = self.Connection.createInstance(service)
        progress(70)
        queries, query = self._getQueries(name, False)
        progress(80)
        composer.setQuery(query.Command)
        print("MergerModel._getQueryComposer() %s - %s" % (query.UpdateTableName, query.Command))
        progress(90)
        return composer, queries

    def _getQueries(self, name, create):
        queries = self.getDataSource().getQueryDefinitions()
        querynames = queries.getElementNames()
        query = self._getQuery(queries, name, create)
        return querynames, query

    def _getQuery(self, queries, name, create):
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

    def _getQueryCommand(self, format):
        query = 'SELECT "%(Column)s" FROM "%(Table)s" ORDER BY "%(Column)s"' % format
        return query

    def _initRowSet(self, datasource, table, filter):
        self._address.DataSourceName = datasource
        self._recipient.DataSourceName = datasource
        column = table.getColumns().getElementNames()
        self._address.Command = self._getRowSetCommand(table.Name, column)
        self._recipient.Filter = filter
        #self._setRowSetOrder()

    def _getRowSetCommand(self, table, column):
        query = 'SELECT %s FROM "%s"' % (column, table)
        return query

# Procedures called internally
    def _getRowSet(self):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.CommandType = COMMAND
        rowset.FetchSize = g_fetchsize
        return rowset






    def getCurrentDocument(self):
        return self._doc

    def getGridColumn1(self, width):
        return self._column1.getColumnModel(width)

    def getGridColumn2(self, width):
        return self._column2.getColumnModel(width)





    def setRowSet(self, *args):
        self.DataSource.setRowSet(*args)

    def executeRecipient(self, *args):
        self.DataSource.executeRecipient(*args)

    def executeAddress(self, *args):
        self.DataSource.executeAddress(*args)

    def executeRowSet(self, *args):
        self.DataSource.executeRowSet(*args)



    def setDocumentRecord(self, index):
        try:
            dispatch = None
            frame = self._doc.getCurrentController().Frame
            flag = uno.getConstantByName('com.sun.star.frame.FrameSearchFlag.SELF')
            if self._doc.supportsService('com.sun.star.text.TextDocument'):
                url = getUrl(self._ctx, '.uno:DataSourceBrowser/InsertContent')
                dispatch = frame.queryDispatch(url, '_self', flag)
            elif self._doc.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
                url = getUrl(self._ctx, '.uno:DataSourceBrowser/InsertColumns')
                dispatch = frame.queryDispatch(url, '_self', flag)
            if dispatch is not None:
                dispatch.dispatch(url, self._getDataDescriptor(index + 1))
        except Exception as e:
            print("PageModel._setDocumentRecord() ERROR: %s - %s" % (e, traceback.print_exc()))

# MergerModel StringResource methods
    def getPageStep(self, id):
        resource = self._getPageStep(id)
        step = self.resolveString(resource)
        return step

    def getPageTitle(self, id):
        resource = self._getPageTitle(id)
        title = self.resolveString(resource)
        return title

# MergerModel StringResource private methods
    def _getPageStep(self, id):
        return 'MergerPage%s.Step' % id

    def _getPageTitle(self, pageid):
        return 'MergerPage%s.Title' % pageid

# MergerModel private methods
    def _getQueryNames(self):
        names = []
        title = self._doc.DocumentProperties.Title.replace(' ', '_')
        if title != '':
            names.append(title)
        template = self._doc.DocumentProperties.TemplateName.replace(' ', '_')
        if template != '':
            names.append(template)
        names.append(g_extension)
        return tuple(names)

    def _getDataDescriptor(self, row):
        descriptor = []
        direct = uno.Enum('com.sun.star.beans.PropertyState', 'DIRECT_VALUE')
        recipient = self.DataSource.DataBase.getRecipient()
        connection = recipient.ActiveConnection
        descriptor.append(PropertyValue('DataSourceName', -1, recipient.DataSourceName, direct))
        descriptor.append(PropertyValue('ActiveConnection', -1, connection, direct))
        descriptor.append(PropertyValue('Command', -1, recipient.Command, direct))
        descriptor.append(PropertyValue('CommandType', -1, recipient.CommandType, direct))
        descriptor.append(PropertyValue('Cursor', -1, recipient, direct))
        descriptor.append(PropertyValue('Selection', -1, (row, ), direct))
        descriptor.append(PropertyValue('BookmarkSelection', -1, False, direct))
        return tuple(descriptor)

    def _getDocumentName(self):
        url = None
        location = self._doc.getLocation()
        if location:
            url = getUrl(self._ctx, location)
        return None if url is None else url.Name
