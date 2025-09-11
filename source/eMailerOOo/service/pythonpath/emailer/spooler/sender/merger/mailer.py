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

from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.text.MailMergeType import FILE

from ....unotool import createService
from ....unotool import getLastNamedParts
from ....unotool import getNamedValueSet
from ....unotool import getPropertyValueSet

from ....helper import getDocumentFilter
from ....helper import hasDocumentFilter

import traceback


class Mailer():
    def __init__(self, ctx, connection, datasource, table, export):
        self._job = self._getJob(ctx, connection, datasource, table, export)
        self._fields = None

    def merge(self, *args):
        raise NotImplementedError('Need to be implemented!')

    def execute(self, sf, descriptor, temp):
        if sf.exists(temp):
            sf.kill(temp)
        self._job.execute(descriptor)

    def getDescriptor(self, document, folder, filter, selection):
        print("Mailer.getDescriptor() 1")
        name, extension = getLastNamedParts(document.Title)
        descriptor = {'DocumentURL': document.getLocation(),
                      'OutputURL': folder,
                      'FileNamePrefix': name,
                      'Filter': selection}
        temp = '%s/%s0' % (folder, name)
        if filter and hasDocumentFilter(document, filter):
            descriptor['SaveFilter'] = getDocumentFilter(document, filter)
            temp += '.' + filter
        elif extension:
            temp += '.' + extension
        print("Mailer.getDescriptor() 2 temp: %s" % temp)
        return getNamedValueSet(descriptor), temp

    def hasInvalidFields(self, document, connection, datasource, table):
        self._fields = self._getInvalidFields(document, connection, datasource, table)
        return self._fields is not None

    def getInvalidFields(self):
        return self._fields

    def dispose(self):
       self._job.dispose()

    def _getJob(self, ctx, connection, datasource, table, export):
        service = 'com.sun.star.text.MailMerge'
        job = createService(ctx, service)
        job.OutputType = FILE
        job.SaveAsSingleFile = True
        job.ActiveConnection = connection
        job.DataSourceName = datasource
        job.Command = table
        job.CommandType = TABLE
        if export:
            job.SaveFilterData = export.getDescriptor()
        return job

    def _getInvalidFields(self, document, connection, datasource, table):
        fields = None
        columns = connection.getTables().getByName(table).getColumns().getElementNames()
        masters = document.getTextFieldMasters()
        for name in masters.getElementNames():
            field = masters.getByName(name)
            fields = self._getMissingFields(field, document.Title, datasource, table, columns)
            if fields:
                break
        return fields

    def _getMissingFields(self, field, title, datasource, table, columns):
        fields = name = None
        properties = field.getPropertySetInfo()
        if self._checkProperty(field, properties, 'DataBaseName', datasource):
            name = 'DataBaseName'
        elif self._checkProperty(field, properties, 'DataTableName', table):
            name = 'DataTableName'
        elif self._checkProperties(field, properties, 'DataColumnName', columns):
            name = 'DataColumnName'
        if name:
            fields = (title, name, field.getPropertyValue(name), datasource)
        return fields

    def _checkProperty(self, field, properties, name, value):
        return properties.hasPropertyByName(name) and field.getPropertyValue(name) != value

    def _checkProperties(self, field, properties, name, values):
        return properties.hasPropertyByName(name) and field.getPropertyValue(name) not in values

