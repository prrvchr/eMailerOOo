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

from com.sun.star.sdb.CommandType import TABLE

from ..export import PdfExport

from ..runner import CancelException
from ..runner import Task

from ..unotool import executeFrameDispatch
from ..unotool import getConfiguration
from ..unotool import getDispatcher
from ..unotool import getDocument
from ..unotool import getLastNamedParts
from ..unotool import getPropertyValue
from ..unotool import getPropertyValueSet
from ..unotool import getSimpleFile
from ..unotool import getStringResource
from ..unotool import getTempFile
from ..unotool import getUriFactory

from ..dbtool import getObjectSequenceFromResult

from ..helper import executeMerge
from ..helper import getConnection
from ..helper import getDataDescriptor
from ..helper import getDataSource
from ..helper import getInvalidFields
from ..helper import getMailUser
from ..helper import getRowSet
from ..helper import saveDocumentTo
from ..helper import tryCorrectFields

from ..configuration import g_identifier
from ..configuration import g_resource
from ..configuration import g_sep

import traceback


class Job(Task):
    def __init__(self, ctx, callback, /, connection=None, document=None, parent=None):
        super().__init__(ctx, 'JobRunner', callback)
        self._ctx = ctx
        self._connection = connection
        self._parent = parent
        self._ownership = True
        self._documents = {}
        self._tmps = {}
        self._urls = {}
        self._sf = None
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
        self._resources = {}
        self._results = {}

    def callback(self, success=True, error=None):
        self._callback(success)

    def stepInitJob(self, ctx):
        self._sf = getSimpleFile(ctx)

    def stepGetJobs(self, jobs=()):
        raise NotImplementedError('Need to be implemented!')

    def stepGetUsers(self, ctx):
        users = {}
        for batch in self._batchs:
            users[batch.Batch] = getMailUser(ctx, batch.Sender)
        self._users = users

    def stepGetStepCount(self):
        raise NotImplementedError('Need to be implemented!')

    def stepInitBatch(self):
        self._close()

    def stepInitDataSource(self, ctx, datasource):
        code, datasource, location = getDataSource(ctx, datasource)
        if code:
            raise Exception()
        self._datasource = datasource

    def stepInitConnection(self, ctx):
        self._connection = getConnection(ctx, self._datasource, self._parent, self._datasource.Name)

    def stepInitRowSet(self, ctx, datasource, table):
        self._rowset = getRowSet(ctx, self._connection, datasource, table)

    def stepInitDocument(self, ctx, document, num):
        self._tmps[num] = getTempFile(ctx).Uri
        self._documents[num] = getDocument(ctx, document, True, True)

    def stepValidateFields(self, table, num):
        document = self._documents[num]
        columns = self._connection.getTables().getByName(table).getColumns().getElementNames()
        fields = getInvalidFields(self._datasource.Name, table, columns, document)
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

    def stepInitDataDescriptor(self, datasource, table):
        self._descriptor = getDataDescriptor(self._connection, datasource, table)

    def stepSetRowsetFilter(self, filter):
        self._rowset.ApplyFilter = False
        self._rowset.Filter = filter
        self._rowset.ApplyFilter = True
        self._rowset.execute()
        self._descriptor['Cursor'] = self._rowset.createResultSet()

    def stepInitPdf(self, ctx):
        if self._pdfExport is None:
            self._pdfExport = PdfExport(ctx)

    def stepMergeDocument(self, ctx, num):
        document = self._documents[num]
        executeMerge(ctx, document, getPropertyValueSet(self._descriptor))

    def stepExportDocument(self, num, export, title):
        document = self._documents[num]
        self._urls[num] = saveDocumentTo(document, self._tmps[num], title, export, self._pdfExport)

    def stepSendThreadMessage(self, ctx, user, batch):
        raise NotImplementedError('Need to be implemented!')

    def stepConnectSmptServer(self, ctx, user):
        raise NotImplementedError('Need to be implemented!')

    def stepBuildMail(self, ctx, user, batch, job, recipient, merge, title):
        raise NotImplementedError('Need to be implemented!')

    def stepSendMail(self, ctx, user, batch, job, recipient, merge):
        raise NotImplementedError('Need to be implemented!')

    def stepFinalize(self):
        self._close()

    def _getStepInitJob(self, ctx, name):
        return 10201, (name, ), self.stepInitJob, ctx

    def _getStepGetJobs(self, jobs=()):
        return 10301, (), self.stepGetJobs, jobs

    def _getStepGetUsers(self, ctx):
        return 10401, (), self.stepGetUsers, ctx

    def _getStepGetStepCount(self):
        return 10501, (), self.stepGetStepCount

    def _getStepInitBatch(self, batch):
        return 10601, (batch, ), self.stepInitBatch

    def _getStepInitDataSource(self, ctx, batch, datasource):
        return 10701, (batch, datasource), self.stepInitDataSource, ctx, datasource

    def _getStepInitConnection(self, ctx, batch):
        return 10801, (batch, ), self.stepInitConnection, ctx

    def _getStepInitRowSet(self, ctx, batch, datasource, table):
        return 10901, (batch, datasource, table), self.stepInitRowSet, ctx, datasource, table

    def _getStepInitDocument(self, ctx, batch, url):
        return 11001, (batch, url), self.stepInitDocument, ctx, url, 0

    def _getStepValidateDocumentFields(self, batch, table, title):
        return 11101, (batch, title), self.stepValidateFields, table, 0

    def _getStepInitDataDescriptor(self, batch, datasource, table):
        return 11201, (batch, table), self.stepInitDataDescriptor, datasource, table

    def _getStepSetRowsetFilter(self, batch, job, filter):
        return 11301, (batch, job, filter), self.stepSetRowsetFilter, filter

    def _getStepMergeDocument(self, ctx, batch, job, title):
        return 11401, (batch, job, title), self.stepMergeDocument, ctx, 0

    def _getStepExportDocument(self, batch, job, export, title):
        return 11501, (batch, job, title), self.stepExportDocument, 0, export, title

    def _getStepInitPdf(self, ctx, batch):
        return 11601, (batch, ), self.stepInitPdf, ctx

    def _getStepInitAttachement(self, ctx, batch, url, num):
        return 11701, (batch, url), self.stepInitDocument, ctx, url, num

    def _getStepValidateAttachmentFields(self, batch, table, num, title):
        return 11801, (batch, title), self.stepValidateFields, table, num

    def _getStepMergeAttachment(self, ctx, batch, job, num, title):
        return 11901, (batch, job, title), self.stepMergeDocument, ctx, num

    def _getStepExportAttachment(self, batch, job, num, export, title):
        return 12001, (batch, job, title), self.stepExportDocument, num, export, title

    def _getStepSendThreadMessage(self, ctx, user, batch):
        return 12101, (batch.Batch, batch.Sender), self.stepSendThreadMessage, ctx, user, batch

    def _getStepConnectSmptServer(self, ctx, user, batch):
        return 12201, (batch.Batch, batch.Sender), self.stepConnectSmptServer, ctx, user

    def _getStepBuildMail(self, ctx, user, batch, job, recipient, merge, title):
        return 12301, (batch.Batch, job, recipient), self.stepBuildMail, ctx, user, batch, job, recipient, merge, title

    def _getStepSendMail(self, ctx, user, batch, job, recipient, merge):
        return 12401, (batch.Batch, job, recipient), self.stepSendMail, ctx, user, batch, job, recipient, merge

    def _getStepFinalize(self):
        return 12501, (), self.stepFinalize

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

        if self._tmps:
            try:
                for tmp in self._tmps.values():
                    if self._sf.exists(tmp) and self._sf.isFolder(tmp):
                        self._sf.kill(tmp)
            except Exception:
                print("Erreur Nettoyage Inter-Batch : %s" % traceback.format_exc())
        self._tmps.clear()
        self._urls.clear()

