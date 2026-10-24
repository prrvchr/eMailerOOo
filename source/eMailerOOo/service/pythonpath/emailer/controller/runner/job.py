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

from ...export import PdfExport

from ...runner import CancelException
from ...runner import Task

from ...unotool import getConfiguration
from ...unotool import getDispatcher
from ...unotool import getDocument
from ...unotool import getLastNamedParts
from ...unotool import getPropertyValue
from ...unotool import getPropertyValueSet
from ...unotool import getStringResource
from ...unotool import getTempFile
from ...unotool import getUriFactory

from ...dbtool import getObjectSequenceFromResult

from ...helper import getConnection
from ...helper import getDataSource
from ...helper import getInvalidFields
from ...helper import getMailUser
from ...helper import getRowSet
from ...helper import saveDocumentTo
from ...helper import tryCorrectFields

from ...configuration import g_identifier
from ...configuration import g_resource

import traceback
from emailer.runner.error import RunnerException


class Job(Task):
    def __init__(self, ctx, callback, /, connection=None, document=None, parent=None):
        super().__init__()
        self._ctx = ctx
        self._callback = callback
        self._connection = connection
        self._parent = parent
        self._ownership = True
        self._documents = {}
        self._urls = {}
        self._datasource = None
        self._rowset = None
        self._decriptor = None
        self._attachments = {}
        self._server = None
        self._pdfExport = None
        self._dispatcher = None
        self._descriptor = None
        config = getConfiguration(ctx, g_identifier)
        self._mode = config.getByName('ErrorMode')
        self.delay = max(0.01, config.getByName('RunDelay'))
        self.resolver = getStringResource(ctx, g_identifier, g_resource, 'JobRunner')
        self.label1 = self.resolver.resolveString(10101)
        self.label2 = self.resolver.resolveString(10102)
        self.label3 = self.resolver.resolveString(10103)
        self._results = {}

    def callback(self, success=True, error=None):
        self._callback(success)

    def stepGetJobs(self):
        raise NotImplementedError('Need to be implemented!')

    def stepGetUsers(self):
        users = {}
        for batch in self._batchs:
            users[batch.Batch] = getMailUser(self._ctx, batch.Sender)
        self._users = users

    def stepGetStepCount(self):
        raise NotImplementedError('Need to be implemented!')

    def stepBuildStep(self, ctx):
        raise NotImplementedError('Need to be implemented!')

    def stepInitBatch(self):
        self._close()

    def stepInitDataSource(self, ctx, datasource):
        code, datasource, location = getDataSource(ctx, datasource)
        if code:
            raise Exception()
        self._datasource = datasource

    def stepInitConnection(self, ctx):
        self._connection = getConnection(ctx, self._datasource, self._parent, name)

    def stepInitRowSet(self, ctx, datasource, table):
        self._rowset = getRowSet(ctx, self._connection, datasource, table)

    def stepInitDispatcher(self, ctx):
        descriptor = getPropertyValueSet({'ActiveConnection': self._connection,
                                          'ResultSet': self._rowset,
                                          'BookmarkSelection': True,
                                          'Selection': (1, )})
        self._descriptor = getPropertyValue('DatabaseDescriptor', descriptor)
        self._dispatcher = getDispatcher(ctx)

    def stepInitDocument(self, ctx, document, num):
        folder, name = getLastNamedParts(getTempFile(ctx).Uri, '/')
        self._urls[num] = folder
        self._documents[num] = getDocument(ctx, document)

    def stepValidateFields(self, document):
        columns = self._connection.getTables().getByName(table).getColumns().getElementNames()
        fields = getInvalidFields(self._datasource.Name, self._table, columns, document)
        if not fields:
            return

        if self._mode == 0:
            raise CancelException()

        elif self._mode == 1:
            self.step += self.skipSteps() 

        else:
            success = True
            for infos in fields:
                if not tryCorrectFields(infos, datasource, table, columns):
                    success = False
                    break
            if success:
                document.getTextFields().refresh()
            else:
                self.step += self.skipSteps()

    def stepInitMerge(self, job, filter):
        self._rowset.ApplyFilter = False
        self._rowset.Filter = filter
        self._rowset.ApplyFilter = True
        self._rowset.execute()

    def stepInitMergeAttachment(self, job, num, filter, url, pdf):
        pass

    def stepInitPdf(self, ctx):
        if self._pdfExport is None:
            self._pdfExport = PdfExport(ctx)

    def stepMergeDocument(self, job, num, filter):
        document = self._documents[num]
        frame = document.getCurrentController().getFrame()
        if document.supportsService('com.sun.star.text.TextDocument'):
            url = '.uno:DataSourceBrowser/InsertContent'
        elif document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
            url = '.uno:DataSourceBrowser/InsertColumns'
        else:
            url = '.uno:DataSourceBrowser/DataToFields'
        self._dispatcher.executeDispatch(frame, url, '', 0, self._descriptor)

    def stepExportDocument(self, job, num, export):
        document = self._documents[num]
        path = self._urls[num]
        url, name = saveDocumentTo(document, path, export, self._pdfExport)
        self._urls[num] = url + '/' + name

    def stepBuildThreadMessage(self, ctx, batch):
        raise NotImplementedError('Need to be implemented!')

    def stepBuildMail(self, job, sender, recipient, subject):
        pass

    def stepAddMailAttachement(self, job, num, url):
        pass

    def stepSendMail(self, job, recipient):
        pass

    def stepFinalize(self):
        self._close()

    def _getStepGetJobs(self):
        return 10201, (), self.stepGetJobs

    def _getStepGetUsers(self):
        return 10301, (), self.stepGetUsers

    def _getStepGetStepCount(self):
        return 10401, (), self.stepGetStepCount

    def _getStepInitBatch(self):
        return 10501, (), self.stepInitBatch

    def _getStepInitDataSource(self, ctx, datasource):
        return 10601, (datasource.Name, ), self.stepInitDataSource, ctx, datasource

    def _getStepInitConnection(self, ctx):
        return 10701, (), self.stepInitConnection, ctx

    def _getStepInitRowSet(self, ctx, datasource, table):
        args = (datasource, table)
        return 10801, args, self.stepInitRowSet, ctx, datasource, table

    def _getStepInitDispatcher(self, ctx):
        return 10901, (), self.stepInitDispatcher, ctx

    def _getStepInitDocument(self, ctx, url):
        args = (url, )
        return 11001, args, self.stepInitDocument, ctx, url, 0

    def _getStepValidateDocumentFields(self, job):
        document = self._documents[0]
        args = (job, document.Title)
        return 11101, args, self.stepValidateFields, document

    def _getStepInitMerge(self, job, filter):
        args = (url, )
        return 11201, args, self.stepInitMerge, job, filter

    def _getStepInitAttachement(self, job, url, num):
        args = (job, url)
        return 11301, args, self.stepInitDocument, url, num

    def _getStepValidateAttachmentFields(self, job, num):
        document = self._documents[num]
        args = (job, document.Title)
        return 11401, args, self.stepValidateFields, document

    def _getStepInitMergeAttachment(self, job, num, url, export):
        return 11501, (), self.stepInitMergeAttachment, job, num, url, export

    def _getStepInitPdf(self, job, num, url, export):
        return 11601, (), self.stepInitPdf, job, num, url, export

    def _getStepMergeDocument(self, job, filter):
        args = (job, self._documents[0].Title)
        return 11701, args, self.stepMergeDocument, job, 0, filter

    def _getStepMergeAttachment(self, job, num, filter):
        args = (job, self._documents[num].Title)
        return 11801, args, self.stepMergeDocument, job, num, filter

    def _getStepExportDocument(self, job, export):
        args = (job, self._documents[0].Title)
        return 11901, args, self.stepExportDocument, job, 0, export

    def _getStepExportAttachment(self, job, num, export):
        args = (job, self._documents[num].Title)
        return 12001, args, self.stepExportDocument, job, num, export

    def _getStepBuildThreadMessage(self, ctx, batch):
        return 12101, (batch.BatchId, ), self.stepBuildThreadMessage, ctx, batch

    def _getStepBuildMail(self, job, sender, recipient, subject):
        return 12201, (), self.stepBuildMail, job, sender, recipient, subject

    def _getStepAddMailAttachement(self, job, num):
        return 12301, (), self.stepAddMailAttachement, job, num

    def _getStepSendMail(self, job, recipient):
        return 12401, (), self.stepSendMail, job, recipient

    def _getStepFinalize(self):
        return 12501, (), self.stepFinalize

    def _buildSteps(self, steps):
        steps.append(self._getStepFinalize())
        size = len(self._steps)
        self._steps.extend(steps)
        self.steps = iter(self._steps)
        for i in range(size):
            next(self.steps)
        self.total = len(self._steps)

    def _close(self):
        if self._server:
            try: self._server.disconnect()
            except: pass
        self._server = None

        if self._rowset:
            try: self._rowset.close()
            except: pass
        self._rowset = None

        if self._ownership and self._connection:
            try: self._connection.close()
            except: pass
        self._connection = None

        if self._datasource:
            try: self._datasource.dispose()
            except: pass
        self._datasource = None

        for document in self._documents.values():
            document.close(True)
        self._documents.clear()

