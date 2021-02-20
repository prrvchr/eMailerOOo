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

from com.sun.star.beans import PropertyValue

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import getDesktop
from unolib import getUrl
from unolib import createService
from unolib import getStringResource

from .datasource import DataSource

from smtpserver import ColumnModel

from smtpserver import g_identifier
from smtpserver import g_extension

from smtpserver import logMessage
from smtpserver import getMessage

import traceback


class MergerModel(unohelper.Base):
    def __init__(self, ctx, email=None):
        self._ctx = ctx
        self._modified = False
        self._stringResource = getStringResource(ctx, g_identifier, g_extension)
        self._doc = getDesktop(ctx).CurrentComponent
        self._column1 = ColumnModel(ctx)
        self._column2 = ColumnModel(ctx)
        self.DataSource = DataSource(ctx)

    def getCurrentDocument(self):
        return self._doc

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

    def getGridColumn1(self, width):
        return self._column1.getColumnModel(width)

    def getGridColumn2(self, width):
        return self._column2.getColumnModel(width)

    def setDataSource(self, progress, datasource, callBack):
        print("PageModel.setDataSource() 1")
        names = self._getQueryNames()
        self.DataSource.setDataSource(progress, datasource, names, callBack)

    def getTableColumns(self, table):
        return self.DataSource.DataBase.getTableColumns(table)

    def setRowSet(self, *args):
        self.DataSource.setRowSet(*args)

    def executeRecipient(self, *args):
        self.DataSource.executeRecipient(*args)

    def executeAddress(self, *args):
        self.DataSource.executeAddress(*args)

    def executeRowSet(self, *args):
        self.DataSource.executeRowSet(*args)

    def getDocumentDataSource(self):
        datasource = ''
        setting = 'com.sun.star.document.Settings'
        if self._doc.supportsService('com.sun.star.text.TextDocument'):
            datasource = self._doc.createInstance(setting).CurrentDatabaseDataSource
        return datasource

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

# MergerModel StringResource private methods
    def _getPageStep(self, id):
        return 'MergerPage%s.Step' % id

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
