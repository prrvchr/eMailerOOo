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

from .batch import Batch

from .attachment import Attachment

from ....dbtool import getRowDict
from ....dbtool import getRowValue

from ....unotool import getLastNamedParts
from ....unotool import getTempFile

from ....helper import executeExport
from ....helper import executeMerge
from ....helper import getConnection
from ....helper import getDataSource
from ....helper import getFilteredRowSet
from ....helper import getJob
from ....helper import getRowSet
from ....helper import parseUriFragment

import json
import traceback


class Mail(Batch):
    def __init__(self, ctx, cancel, resolver, sf, uf, export, batch, filter='html'):
        super().__init__(sf, batch, batch.Document, getTempFile(ctx).Uri, batch.Merge, filter)
        folder, name = getLastNamedParts(batch.Document, '/')
        self._title = name
        if sf.exists(batch.Document):
            self._exists = True
            self._job, self._datasource, self._rowset, self._rows = self._getJob(ctx, resolver, batch, export)
        else:
            self._exists = False
            self._job = self._datasource = self._rowset = self._rows = None
        self._attachments, self._missings = self._parseAttachments(cancel, sf, uf, batch)
        self._export = export
        self._fields = None

    @property
    def Batch(self):
        return self._batch

    @property
    def Subject(self):
        return self._batch.Subject

    @property
    def Sender(self):
        return self._batch.Sender

    @property
    def Query(self):
        return self._batch.Query

    @property
    def Document(self):
        return self._batch.Document

    @property
    def Title(self):
        return self._title

    @property
    def Attachments(self):
        return self._attachments

    @property
    def Jobs(self):
        return self._batch.Jobs

    @property
    def JobCount(self):
        return len(self._batch.Jobs)

    def hasFile(self):
        return self._exists

    def hasAttachments(self):
        return len(self._missings) == 0

    def getMissingAttachments(self):
        return ', '.join(self._missings)

    def hasInvalidFields(self):
        return self._fields is not None

    def getInvalidFields(self):
        return self._fields

    def getJob(self):
        return self._job

    def execute(self, cancel, progress):
        if not cancel.isSet():
            if self._job:
                print("Mail.execute() 1")
                self._fields = executeMerge(self._job, self)
                print("Mail.execute() 2")
            else:
                executeExport(self._ctx, self, self._export)
            if not self.hasInvalidFields() and not cancel.isSet():
                self._fields = self._proccessAttachments(cancel, progress)

    def close(self):
        print("Mail.close() 1 BatchId: %s" % self.BatchId)
        # XXX: If we want to avoid a memory dump when exiting LibreOffice,
        # XXX: it is imperative to close / dispose all these references.
        if self._job:
            result = self._job.getPropertyValue('ResultSet')
            if result:
                result.close()
        if self._rowset:
            self._rowset.close()
        if self._job:
            connection = self._job.getPropertyValue('ActiveConnection')
            if connection:
                connection.close()
            self._job.dispose()
        print("Mail.close() 2")

    def getRecipient(self, job):
        if self._merge:
            recipient = self._rows[job][1]
        else:
            index = self._jobs.index(job)
            recipient = self._recipients[index]
        return recipient

    def getSubject(self, job):
        if self._merge:
            subject = self._rows[job][2]
        else:
            subject = self._batch.Subject
        return subject

    def getDataSource(self):
        return self._datasource

    def hasDataSource(self):
        return self._datasource is not None

# Private methods
    def _getJob(self, ctx, resolver, batch, export):
        rows = {}
        job = datasource = rowset = None
        if batch.Merge:
            datasource = getDataSource(ctx, batch.DataSource, resolver, 1531)
            if datasource:
                connection = getConnection(ctx, datasource)
                if connection:
                    rs = getRowSet(ctx, connection, batch.DataSource, batch.Table)
                    rowset = getFilteredRowSet(rs, self._getFilters(batch.Filters))
                    result = rowset.createResultSet()
                    rows = self._getJobRows(batch, result)
                    job = getJob(ctx, connection, batch.DataSource, batch.Table, result, export)
        return job, datasource, rowset, rows

    def _getFilters(self, filters):
        return ' OR '.join(filters)

    def _getJobRows(self, batch, result):
        metadata = result.getMetaData()
        count = metadata.getColumnCount()
        emails = self._getEmailIndexes(batch, metadata, count)
        return self._getJobIdentifiers(batch, result, metadata, count, emails)

    def _getEmailIndexes(self, batch, metadata, count):
        indexes = []
        for address in batch.Addresses:
            for i in range(1, count + 1):
                if metadata.getColumnLabel(i) == address:
                    indexes.append(i)
        return tuple(indexes)

    def _getJobIdentifiers(self, batch, result, metadata, count, data=None):
        identifiers = {}
        indexes = self._getIdentifierIndexes(batch, metadata, count)
        identifier = 0
        while result.next():
            print("Job._getJobIdentifiers() identifier: %s" % identifier)
            predicate = self._getRowPredicate(result, metadata, indexes)
            if predicate in batch.Predicates:
                index = batch.Predicates.index(predicate)
                job = batch.Jobs[index]
                identifiers[job] = self._getRowData(identifier, result, metadata, data)
            identifier += 1
        return identifiers

    def _getIdentifierIndexes(self, batch, metadata, count):
        indexes = []
        for identifier in batch.Identifiers:
            for i in range(1, count + 1):
                if metadata.getColumnLabel(i) == identifier:
                    indexes.append(i)
        return tuple(indexes)

    def _getRowPredicate(self, result, metadata, indexes):
        predicates = []
        for i in indexes:
            value = getRowValue(result, metadata.getColumnType(i), i)
            predicates.append(value)
        return json.dumps(predicates)

    def _getRowData(self, identifier, result, metadata, data=None):
        recipient = self._getRowRecipient(result, metadata, data)
        subject = self._getSubject(result)
        return identifier, recipient, subject

    def _getRowRecipient(self, result, metadata, indexes):
        recipient = None
        for i in indexes:
            value = getRowValue(result, metadata.getColumnType(i), i)
            if value is not None:
                recipient = value
                break
        return recipient

    def _getSubject(self, result):
        subject = self._batch.Subject
        fields = getRowDict(result, '')
        try:
            subject = subject.format(**fields)
        except ValueError as e:
            pass
        return subject

    def _parseAttachments(self, cancel, sf, uf, batch):
        attachments = []
        missings = []
        for attachment in batch.Attachments:
            if cancel.isSet():
                break
            url, merge, filter = parseUriFragment(uf, attachment, batch.Merge)
            if sf.exists(url):
                attachments.append(Attachment(sf, batch, self.Folder, self._rows, url, merge, filter))
            else:
                missings.append(url)
        return attachments, missings

    def _proccessAttachments(self, cancel, progress):
        fields = None
        for attachment in self.Attachments:
            if cancel.isSet():
                break
            if self._job and attachment.Merge:
                print("Mail._proccessAttachments() 1")
                fields = executeMerge(self._job, attachment)
                print("Mail._proccessAttachments() 2")
                if fields is not None:
                    break
            elif attachment.Filter:
                executeExport(self._ctx, attachment, self._export)
        return fields

