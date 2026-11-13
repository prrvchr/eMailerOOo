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

from ..export import PdfExport

from ..unotool import Transferable

from ..runner import CancelException
from ..runner import Task

from ..unotool import getDocument
from ..unotool import getLastNamedParts
from ..unotool import getPathSettings
from ..unotool import getPropertyValueSet
from ..unotool import getSimpleFile
from ..unotool import getTempFile
from ..unotool import getUriFactory
from ..unotool import getUriTranslator

from ..helper import executeMerge
from ..helper import getConnection
from ..helper import getDataDescriptor
from ..helper import getDataSource
from ..helper import getInvalidFields
from ..helper import getRowSet
from ..helper import saveDocumentTo
from ..helper import tryCorrectFields

from ..configuration import g_sep

import time
import traceback


class Job(Task):
    def __init__(self, ctx, callback, parent):
        super().__init__(ctx, 'JobRunner', callback)
        self._ctx = ctx
        self._connection = None
        self._parent = parent
        self._documents = {}
        self._work = None
        self._tmps = {}
        self._urls = {}
        self._datasource = None
        self._rowset = None
        self._decriptor = None
        self._attachments = {}
        self._server = None
        self._pdfExport = None
        self._dispatcher = None
        self._descriptor = None
        self._mode = 0
        self._sf = None
        self._factory = None
        self._transferable = None
        self._resources = {}
        self._results = []

    def callback(self, success=True, error=None):
        self._callback(success, self._results, error)

    def getStepFinalize(self):
        return self._getStepFinalize()

    def stepInitJob(self, ctx):
        self._work = getPathSettings(ctx).Work
        self._sf = getSimpleFile(ctx)
        self._factory = getUriFactory(ctx)
        self._translator = getUriTranslator(ctx)
        self._transferable = Transferable(ctx)

    def stepGetJobs(self, jobs=()):
        raise NotImplementedError('Need to be implemented!')

    def stepGetUsers(self, ctx):
        raise NotImplementedError('Need to be implemented!')

    def stepGetStepCount(self, *args):
        raise NotImplementedError('Need to be implemented!')

    def stepInitBatch(self):
        self._close()

    def stepInitDataSource(self, ctx, datasource):
        code, datasource, _ = getDataSource(ctx, datasource)
        if code:
            raise Exception()
        self._datasource = datasource

    def stepInitConnection(self, ctx):
        self._connection = getConnection(ctx, self._datasource, self._parent, self._datasource.Name)

    def stepInitRowSet(self, ctx, datasource, table):
        self._rowset = getRowSet(ctx, self._connection, datasource, table)

    def stepInitDataDescriptor(self, datasource, table):
        self._descriptor = getDataDescriptor(self._connection, datasource, table)

    def stepInitPdf(self, ctx):
        if self._pdfExport is None:
            self._pdfExport = PdfExport(ctx)

    def stepInitDocument(self, ctx, document, num):
        self._tmps[num] = getTempFile(ctx).Uri
        self._documents[num] = getDocument(ctx, document, True, True)

    def stepValidateFields(self, datasource, table, num):
        document = self._documents[num]
        columns = self._connection.getTables().getByName(table).getColumns().getElementNames()
        fields = getInvalidFields(datasource, table, columns, document)
        if not fields:
            return

        if self._mode == 0:
            raise CancelException()

        elif self._mode == 1:
            self.step += self.skipSteps()

        else:
            success = True
            for infos in fields:
                if not tryCorrectFields(datasource, table, columns, *infos):
                    success = False
                    break
            if success:
                document.getTextFields().refresh()
            else:
                self.step += self.skipSteps()

    def stepSetRowsetFilter(self, fltr):
        self._rowset.ApplyFilter = False
        self._rowset.Filter = fltr
        self._rowset.ApplyFilter = True
        self._rowset.execute()
        self._descriptor['Cursor'] = self._rowset.createResultSet()

    def stepMergeDocument(self, ctx, num):
        document = self._documents[num]
        executeMerge(ctx, document, getPropertyValueSet(self._descriptor))

    def stepExportDocument(self, num, export, title):
        document = self._documents[num]
        self._urls[num] = saveDocumentTo(document, self._tmps[num], title, export, self._pdfExport)

    def stepBuildMail(self, ctx, user, recipient, merge, title, batch):
        raise NotImplementedError('Need to be implemented!')

    def stepCloneDocument(self):
        source = self._urls[0]
        _, title = getLastNamedParts(source, g_sep)
        target = self._work + g_sep + title
        if self._sf.exists(target):
            target = self._getTargetUniqueName(*getLastNamedParts(title))
        self._sf.copy(source, target)
        self._results.append(self._translator.translateToInternal(target))

    def stepSendThreadMessage(self, ctx, user, batch):
        raise NotImplementedError('Need to be implemented!')

    def stepConnectSmptServer(self, ctx, user):
        raise NotImplementedError('Need to be implemented!')

    def stepSendMail(self, ctx, user, recipient, merge, batch, job):
        raise NotImplementedError('Need to be implemented!')

    def stepFinalize(self):
        self._close()

    def _getStepInitJob(self, ctx, name):
        return 10201, (name, ), self.stepInitJob, ctx

    def _getStepGetJobs(self, jobs=()):
        return 10301, (), self.stepGetJobs, jobs

    def _getStepGetUsers(self, ctx):
        return 10401, (), self.stepGetUsers, ctx

    def _getStepGetStepCount(self, *args):
        return 10501, (), self.stepGetStepCount, *args

    def _getStepInitBatch(self, batch):
        return 10601, (batch, ), self.stepInitBatch

    def _getStepInitDataSource(self, ctx, datasource, batch):
        return 10701, (batch, datasource), self.stepInitDataSource, ctx, datasource

    def _getStepInitConnection(self, ctx, batch):
        return 10801, (batch, ), self.stepInitConnection, ctx

    def _getStepInitRowSet(self, ctx, datasource, table, batch):
        return 10901, (batch, datasource, table), self.stepInitRowSet, ctx, datasource, table

    def _getStepInitDataDescriptor(self, datasource, table, batch):
        return 11001, (batch, table), self.stepInitDataDescriptor, datasource, table

    def _getStepInitPdf(self, ctx, batch):
        return 11101, (batch, ), self.stepInitPdf, ctx

    def _getStepInitDocument(self, ctx, url, batch):
        return 11201, (batch, url), self.stepInitDocument, ctx, url, 0

    def _getStepInitAttachement(self, ctx, url, num, batch):
        return 11301, (batch, url), self.stepInitDocument, ctx, url, num

    def _getStepValidateDocumentFields(self, datasource, table, title, batch):
        return 11401, (batch, title), self.stepValidateFields, datasource, table, 0

    def _getStepValidateAttachmentFields(self, datasource, table, num, title, batch):
        return 11501, (batch, title), self.stepValidateFields, datasource, table, num

    def _getStepSetRowsetFilter(self, fltr, batch, job):
        return 11601, (batch, job, fltr), self.stepSetRowsetFilter, fltr

    def _getStepMergeDocument(self, ctx, title, batch, job):
        return 11701, (batch, job, title), self.stepMergeDocument, ctx, 0

    def _getStepMergeAttachment(self, ctx, num, title, batch, job):
        return 11801, (batch, job, title), self.stepMergeDocument, ctx, num

    def _getStepExportDocument(self, export, title, batch, job):
        return 11901, (batch, job, title), self.stepExportDocument, 0, export, title

    def _getStepExportAttachment(self, num, export, title, batch, job):
        return 12001, (batch, job, title), self.stepExportDocument, num, export, title

    def _getStepBuildMail(self, ctx, user, recipient, merge, title, batch, job):
        return 12101, (batch.Batch, job, recipient), self.stepBuildMail, ctx, user, recipient, merge, title, batch

    def _getStepCloneDocument(self, title, batch, job):
        return 12201, (batch, job, title), self.stepCloneDocument

    def _getStepSendThreadMessage(self, ctx, user, batch):
        return 12301, (batch.Batch, batch.Sender), self.stepSendThreadMessage, ctx, user, batch

    def _getStepConnectSmptServer(self, ctx, user, batch):
        return 12401, (batch.Batch, batch.Sender), self.stepConnectSmptServer, ctx, user

    def _getStepSendMail(self, ctx, user, recipient, merge, batch, job):
        return 12501, (batch.Batch, job, recipient), self.stepSendMail, ctx, user, recipient, merge, batch, job

    def _getStepFinalize(self):
        return 12601, (), self.stepFinalize

    def _close(self):
        if self._server:
            try:
                self._server.disconnect()
            except Exception:
                pass
        self._server = None

        if self._rowset:
            try:
                self._rowset.close()
            except Exception:
                pass
        self._rowset = None

        if self._connection:
            try:
                self._connection.close()
            except Exception:
                pass
        self._connection = None

        if self._datasource:
            try:
                self._datasource.dispose()
            except Exception:
                pass
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
        time.sleep(0.075)

    def _getTargetUniqueName(self, name, extension):
        i = 1
        target = self._work + g_sep + f'{name}-{i}.{extension}'
        while self._sf.exists(target):
            i += 1
            target = self._work + g_sep + f'{name}-{i}.{extension}'
        return target
