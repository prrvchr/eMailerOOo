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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.uno import Exception as UnoException

from ..mailerlib import MailTransferable

from ..unotool import createService
from ..unotool import executeFrameDispatch
from ..unotool import getDocument
from ..unotool import getInteractionHandler
from ..unotool import getPropertyValueSet
from ..unotool import getSimpleFile
from ..unotool import getTempFile
from ..unotool import getUriFactory

from ..dbtool import getRowDict

from ..mailertool import getMailConfiguration
from ..mailertool import saveTempDocument


from ..configuration import g_fetchsize

import traceback


class Mailer(unohelper.Base):
    def __init__(self, ctx, database, logger, send=False):
        self._ctx = ctx
        self._database = database
        self._logger = logger
        self._send = send
        self._sf = getSimpleFile(ctx)
        self._uf = getUriFactory(ctx)
        self._batch = None
        self._user = None
        self._descriptor = None
        self._url = None
        self._urls = ()
        self._thread = None
        self._rowset = None
        self._statement = None

    @property
    def Merge(self):
        return self._metadata.get('Merge')
    @property
    def Sender(self):
        return self._metadata.get('Sender')
    @property
    def Document(self):
        return self._metadata.get('Document')
    @property
    def DataSource(self):
        return self._metadata.get('DataSource')
    @property
    def Table(self):
        return self._metadata.get('Table')
    @property
    def Query(self):
        return self._metadata.get('Query')
    @property
    def ThreadId(self):
        return self._metadata.get('ThreadId')

    def isNewBatch(self, batch):
        new = self._batch != batch
        if new and self._batch is not None:
            self._dispose()
        return new

    def setBatch(self, batch, metadata, attachments, job, filter):
        self._batch = batch
        self._user = getMailConfiguration(self._ctx, metadata.get('Sender'))
        self._metadata = metadata
        self._checkUrl(self.Document, job, 161)
        self._rowset, self._descriptor = self._getDescriptors()
        self._urls, self._url = self._getUrls(attachments, job, filter)

    def needThreadId(self):
        return self._send and self.Merge and self._user.supportIMAP() and self.ThreadId is None

    def setThreadId(self, connection, batchid, thread):
        self._database.updateMailer(connection, batchid, thread)
        self._metadata['ThreadId'] = thread

    def getSubject(self, merge=True):
        subject = self._metadata.get('Subject')
        if merge and self.Merge:
            fields = self._getSubjectFields()
            try:
                subject = subject.format(**fields)
            except ValueError as e:
                pass
        return subject

    def dispose(self):
        if self._batch is not None:
            self._dispose()

    def merge(self, filter):
        descriptor = self._getFilteredDescriptor(filter)
        self._url.merge(descriptor)

    def getBodyUrl(self):
        return self._url.Main

    def getDocumentTitle(self):
        return self._url.Title

    def getAttachments(self, tag, separator):
        urls = []
        for url in self._urls:
            urls.append(tag % (url.Url, url.Name))
        return separator.join(urls)

    def addAttachments(self, mail, filter):
        if self._hasThreadId():
            mail.ThreadId = self.ThreadId
        for url in self._urls:
            if self.Merge and url.Merge:
                descriptor = self._getFilteredDescriptor(filter)
                url.merge(descriptor)
            mail.addAttachment(self._getAttachment(url))

    def hasAttachments(self):
        return len(self._urls) > 0

    def getUser(self):
        return self._user

# Private Procedures call
    def _getSubjectFields(self):
        fields = {}
        result = self._rowset.createResultSet()
        if result.next():
            fields = getRowDict(result, '')
        result.close()
        return fields

    def _dispose(self):
        if self._descriptor is not None:
            self._descriptor['ActiveConnection'].close()
        self._url.dispose()
        for url in self._urls:
            url.dispose()

    def _hasThreadId(self):
        return self.ThreadId is not None

    def _getAttachment(self, url):
        attachment = uno.createUnoStruct('com.sun.star.mail.MailAttachment')
        attachment.Data = MailTransferable(self._ctx, url.Main, False, True)
        attachment.ReadableName = url.Name
        return attachment

    def _getDescriptors(self):
        rowset = None
        descriptor = None
        if self.Merge:
            service = 'com.sun.star.sdb.DatabaseContext'
            datasource = createService(self._ctx, service).getByName(self.DataSource)
            connection = self._getDataSourceConnection(datasource)
            rowset = self._getRowSet(connection)
            descriptor = {'DataSourceName': self.DataSource,
                          'ActiveConnection': connection,
                          'Command': self.Table,
                          'CommandType': TABLE,
                          'BookmarkSelection': False,
                          'Selection': (1, )}
        return rowset, descriptor

    def _getDataSourceConnection(self, datasource):
        if datasource.IsPasswordRequired:
            handler = getInteractionHandler(self._ctx)
            connection = datasource.getIsolatedConnectionWithCompletion(handler)
        else:
            connection = datasource.getIsolatedConnection('', '')
        return connection

    def _getRowSet(self, connection):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.ActiveConnection = connection
        rowset.DataSourceName = self.DataSource
        rowset.CommandType = TABLE
        rowset.FetchSize = g_fetchsize
        rowset.Command = self.Table
        return rowset

    def _getUrls(self, attachments, job, filter):
        descriptor = self._getFilteredDescriptor(filter) if self.Merge else None
        uri = self._uf.parse(self.Document)
        url = MailUrl(self._ctx, uri, self.Merge, 'html', descriptor)
        urls = self._getMailUrl(attachments, job)
        return urls, url

    def _getFilteredDescriptor(self, filter):
        self._rowset.ApplyFilter = False
        self._rowset.Filter = filter
        self._rowset.ApplyFilter = True
        self._rowset.execute()
        self._descriptor['Cursor'] = self._rowset.createResultSet()
        descriptor = getPropertyValueSet(self._descriptor)
        return descriptor

    def _getMailUrl(self, attachments, job):
        urls = []
        for attachment in attachments:
            url = self._uf.parse(attachment)
            merge, filter = self._getUrlMark(url)
            self._checkUrl(url.UriReference, job, 171)
            urls.append(MailUrl(self._ctx, url, merge, filter))
        return urls

    def _getUrlMark(self, url):
        merge = False
        filter = None
        if url.hasFragment():
            fragment = url.getFragment()
            if fragment.startswith('merge'):
                merge = True
            if fragment.endswith('pdf'):
                filter = 'pdf'
            url.clearFragment()
        return merge, filter

    def _checkUrl(self, url, job, resource):
        if not self._sf.exists(url):
            raise getUnoException(self._logger, self, resource, job, url)


class MailUrl(unohelper.Base):
    def __init__(self, ctx, url, merge, filter=None, descriptor=None):
        self._ctx = ctx
        self._url = url.UriReference
        self._merge = merge
        self._filter = filter
        name = url.getPathSegment(url.PathSegmentCount -1)
        self._name = self._title = name
        self._temp = self._document = None
        if self._isTemp():
            self._temp = getTempFile(ctx).Uri
            self._document = getDocument(ctx, self._url)
            if descriptor is not None:
                self.merge(descriptor)
            elif not self._merge:
                self._name = self._saveTempDocument()

    @property
    def Merge(self):
        return self._merge
    @property
    def Name(self):
        return self._name
    @property
    def Title(self):
        return self._title
    @property
    def Url(self):
        return self._url
    @property
    def Main(self):
        if self._isTemp():
            url = self._temp
        else:
            url = self._url
        return url

    def merge(self, descriptor):
        self._setDocumentRecord(descriptor)
        self._name = self._saveTempDocument()

    def dispose(self):
        if self._isTemp():
            self._document.close(True)
            self._temp = None

# Private Procedures call
    def _isTemp(self):
        return any((self._merge, self._filter))

    def _setDocumentRecord(self, descriptor):
        url = None
        if self._document.supportsService('com.sun.star.text.TextDocument'):
            url = '.uno:DataSourceBrowser/InsertContent'
        elif self._document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
            url = '.uno:DataSourceBrowser/InsertColumns'
        if url is not None:
            frame = self._document.CurrentController.Frame
            executeFrameDispatch(self._ctx, frame, url, descriptor)

    def _saveTempDocument(self):
        return saveTempDocument(self._document, self._temp, self._title, self._filter)


def getUnoException(logger, source, resource, *args):
    e = UnoException()
    e.Message = logger.resolveString(resource, *args)
    e.Context = source
    return e

