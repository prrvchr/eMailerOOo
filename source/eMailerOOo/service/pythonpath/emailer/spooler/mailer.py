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

from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.mail.MailServiceType import IMAP

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.uno import Exception as UnoException

from ..mailerlib import MailTransferable

from ..unotool import createService
from ..unotool import getInteractionHandler
from ..unotool import getPropertyValueSet
from ..unotool import getResourceLocation
from ..unotool import getSimpleFile
from ..unotool import getUriFactory

from ..dbtool import getRowDict

from ..mailertool import getMailConfiguration
from ..mailertool import getMailMessage
from ..mailertool import getMailService
from ..mailertool import getMessageImage

from ..configuration import g_extension
from ..configuration import g_identifier
from ..configuration import g_logourl
from ..configuration import g_fetchsize
from ..configuration import g_logo

from .mailurl import MailUrl

import traceback


class Mailer():
    def __init__(self, ctx, database, logger, send=False):
        self._ctx = ctx
        self._database = database
        self._logger = logger
        self._send = send
        self._server = None
        self._sf = getSimpleFile(ctx)
        self._uf = getUriFactory(ctx)
        self._batch = None
        self._user = None
        self._descriptor = None
        self._url = None
        self._urls = ()
        self._rowset = None
        self._threadid = None
        logo = '%s/%s' % (g_extension, g_logo)
        self._logo = getResourceLocation(ctx, g_identifier, logo)

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

    def dispose(self):
        if self._server is not None:
            self._server.disconnect()
        if self._descriptor is not None:
            self._descriptor['ActiveConnection'].close()
        self._url.dispose()
        for url in self._urls:
            url.dispose()

    def getMail(self, job):
        recipient = self._database.getRecipient(job)
        if self._isNewBatch(recipient.BatchId):
            self._initMailer(job, recipient)
        elif self.Merge:
            self._merge(recipient.Filter)
        body = MailTransferable(self._ctx, self._getBodyUrl(), True, True)
        mail = getMailMessage(self._ctx, self.Sender, recipient.Recipient, self._getSubject(), body)
        if self._hasThread():
            mail.ThreadId = self._threadid
        self._addAttachments(mail, recipient.Filter)
        return mail

    def sendMail(self, mail):
        self._server.sendMailMessage(mail)

# Private Procedures call
    def _initMailer(self, job, recipient):
        self._threadid = None
        self._batch = recipient.BatchId
        self._metadata = self._database.getMailer(self._batch)
        self._user = getMailConfiguration(self._ctx, self.Sender)
        self._checkUrl(self.Document, job, 161)
        self._rowset, self._descriptor = self._getDescriptors()
        self._urls, self._url = self._getUrls(job, recipient.Filter)
        if self._send:
            self._server = self._getMailServer(SMTP)
            if self._needThread():
                self._createThread()

    def _createThread(self):
        server = self._getMailServer(IMAP)
        folder = server.getSentFolder()
        if server.hasFolder(folder):
            subject = self._getSubject(False)
            message = self._getThreadMessage(subject)
            body = MailTransferable(self._ctx, message, True)
            mail = getMailMessage(self._ctx, self.Sender, self.Sender, subject, body)
            server.uploadMessage(folder, mail)
            self._threadid = mail.MessageId
        server.disconnect()

    def _getThreadMessage(self, subject):
        title = self._logger.resolveString(1031, g_extension, self._batch, self.Query)
        message = self._logger.resolveString(1032)
        document = self._logger.resolveString(1033)
        files = self._logger.resolveString(1034)
        if self._hasAttachments():
            tag = '<a href="%s">%s</a>'
            separator = '</li><li>'
            attachments = '<ol><li>%s</li></ol>' % self._getAttachments(tag, separator)
        else:
            attachments = '<p>%s</p>' % self._logger.resolveString(1035)
        logo = getMessageImage(self._ctx, self._logo)
        return '''\
<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
  </head>
  <body>
    <img alt="%s Logo" src="data:image/png;charset=utf-8;base64,%s" src="%s" />
    <h3 style="display:inline;" >&nbsp;%s</h3>
    <p><b>%s:</b>&nbsp;%s</p>
    <p><b>%s:</b>&nbsp;<a href="%s">%s</a></p>
    <p><b>%s:</b></p>
    %s
  </body>
</html>
''' % (g_extension, logo, g_logourl, title, message, subject,
        document, self.Document, self._getDocumentTitle(),
        files, attachments)

    def _getMailServer(self, mailtype):
        domain = self._user.getServerDomain(mailtype)
        server = getMailService(self._ctx, mailtype.value, domain)
        context = self._user.getConnectionContext(mailtype)
        authenticator = self._user.getAuthenticator(mailtype)
        server.connect(context, authenticator)
        return server

    def _isNewBatch(self, batch):
        new = self._batch != batch
        if new and self._batch is not None:
            self.dispose()
        return new

    def _needThread(self):
        return self.Merge and self._user.supportIMAP()

    def _getSubject(self, merge=True):
        subject = self._metadata.get('Subject')
        if merge and self.Merge:
            fields = self._getSubjectFields()
            try:
                subject = subject.format(**fields)
            except ValueError as e:
                pass
        return subject

    def _merge(self, filter):
        descriptor = self._getFilteredDescriptor(filter)
        self._url.merge(descriptor)

    def _getBodyUrl(self):
        return self._url.Main

    def _getDocumentTitle(self):
        return self._url.Title

    def _getAttachments(self, tag, separator):
        urls = []
        for url in self._urls:
            urls.append(tag % (url.Url, url.Name))
        return separator.join(urls)

    def _addAttachments(self, mail, filter):
        for url in self._urls:
            if self.Merge and url.Merge:
                descriptor = self._getFilteredDescriptor(filter)
                url.merge(descriptor)
            mail.addAttachment(self._getAttachment(url))

    def _hasAttachments(self):
        return len(self._urls) > 0

    def _getSubjectFields(self):
        fields = {}
        result = self._rowset.createResultSet()
        if result.next():
            fields = getRowDict(result, '')
        result.close()
        return fields

    def _hasThread(self):
        return self._threadid is not None

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

    def _getUrls(self, job, filter):
        uri = self._uf.parse(self.Document)
        descriptor = self._getFilteredDescriptor(filter) if self.Merge else None
        url = MailUrl(self._ctx, uri, self.Merge, 'html', descriptor)
        urls = self._getMailUrl(job)
        return urls, url

    def _getFilteredDescriptor(self, filter):
        self._rowset.ApplyFilter = False
        self._rowset.Filter = filter
        self._rowset.ApplyFilter = True
        self._rowset.execute()
        self._descriptor['Cursor'] = self._rowset.createResultSet()
        descriptor = getPropertyValueSet(self._descriptor)
        return descriptor

    def _getMailUrl(self, job):
        urls = []
        for attachment in self._database.getAttachments(self._batch):
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


def getUnoException(logger, source, resource, *args):
    e = UnoException()
    e.Message = logger.resolveString(resource, *args)
    e.Context = source
    return e

