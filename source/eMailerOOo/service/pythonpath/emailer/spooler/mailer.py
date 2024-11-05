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

from com.sun.star.mail import MailSpoolerException

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.uno import Exception as UnoException

from ..transferable import Transferable

from ..mailertool import getDataBaseContext

from ..unotool import createService
from ..unotool import getInteractionHandler
from ..unotool import getPropertyValueSet
from ..unotool import getResourceLocation
from ..unotool import getSimpleFile
from ..unotool import getUriFactory

from ..dbtool import getRowDict

from ..mailertool import getMailUser
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
    def __init__(self, ctx, source, database, logger, send=False):
        self._ctx = ctx
        self._source = source
        self._database = database
        self._logger = logger
        self._send = send
        self._server = None
        self._sf = getSimpleFile(ctx)
        self._uf = getUriFactory(ctx)
        self._transferable = Transferable(ctx, logger)
        self._batch = None
        self._user = None
        self._descriptor = None
        self._url = None
        self._urls = ()
        self._rowset = None
        self._threadid = None
        self._reply = False
        self._replyto = ''
        logo = '%s/%s' % (g_extension, g_logo)
        self._logo = getResourceLocation(ctx, g_identifier, logo)
        self._logger.logprb(INFO, 'Mailer', '__init__()', 1501)

    @property
    def _merge(self):
        return self._metadata.get('Merge')

    def dispose(self):
        self._dispose()
        self._logger.logprb(INFO, 'Mailer', 'dispose()', 1561)

    def getMail(self, job):
        recipient = self._database.getRecipient(job)
        if self._isNewBatch(recipient.BatchId):
            self._initMailer(job, recipient)
        elif self._merge:
            self._mergeMail(recipient.Filter)
        body = self._transferable.getByUrl(self._url.Main)
        mail = getMailMessage(self._ctx, self._getSender(), recipient.Recipient, self._getSubject(), body)
        if self._hasThread():
            mail.ThreadId = self._threadid
        if self._reply:
            mail.ReplyToAddress = self._replyto
        self._addAttachments(mail, recipient.Filter)
        return mail

    def sendMail(self, mail):
        self._server.sendMailMessage(mail)

# Private Procedures call
    def _dispose(self):
        if self._server is not None:
            self._server.disconnect()
        if self._descriptor is not None:
            self._descriptor['ActiveConnection'].close()
        if self._url is not None:
            self._url.dispose()
        for url in self._urls:
            url.dispose()

    def _initMailer(self, job, recipient):
        threadid = None
        metadata = self._database.getMailer(recipient.BatchId)
        merge = metadata.get('Merge')
        url = metadata.get('Document')
        self._checkUrl(url, job, 1511)
        self._urls = self._getMailAttachmentUrl(recipient.BatchId, job)
        self._rowset, self._descriptor = self._getDescriptors(merge, metadata)
        self._url = self._getMailUrl(recipient, url, merge)
        sender = metadata.get('Sender')
        user = self._getMailUser(job, sender)
        self._reply = user.useReplyTo()
        self._replyto = user.getReplyToAddress()
        if self._send:
            self._server = self._getMailServer(user, SMTP)
            if self._needThread(user, merge):
                threadid = self._createThread(recipient.BatchId, user, url, sender, metadata)
        self._batch = recipient.BatchId
        self._metadata = metadata
        self._threadid = threadid

    def _createThread(self, batch, user, url, sender, metadata):
        threadid = None
        server = self._getMailServer(user, IMAP)
        folder = server.getSentFolder()
        if server.hasFolder(folder):
            subject = metadata.get('Subject')
            message = self._getThreadMessage(batch, url, subject, metadata.get('Query'))
            body = self._transferable.getByString(message)
            mail = getMailMessage(self._ctx, sender, sender, subject, body)
            server.uploadMessage(folder, mail)
            threadid = mail.MessageId
        server.disconnect()
        return threadid

    def _getThreadMessage(self, batch, url, subject, query):
        title = self._logger.resolveString(1551, g_extension, batch, query)
        message = self._logger.resolveString(1552)
        label = self._logger.resolveString(1553)
        files = self._logger.resolveString(1554)
        if self._hasAttachments():
            tag = '<a href="%s">%s</a>'
            separator = '</li><li>'
            attachments = '<ol><li>%s</li></ol>' % self._getAttachments(tag, separator)
        else:
            attachments = '<p>%s</p>' % self._logger.resolveString(1555)
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
        label, url, self._getDocumentTitle(),
        files, attachments)

    def _getMailUser(self, job, sender):
        user = getMailUser(self._ctx, sender)
        if user is None:
            msg = self._getErrorMessage(1541, job, sender)
            raise MailSpoolerException(msg, self._source, ())
        return user

    def _getMailServer(self, user, mailtype):
        server = getMailService(self._ctx, mailtype.value)
        context = user.getConnectionContext(mailtype)
        authenticator = user.getAuthenticator(mailtype)
        server.connect(context, authenticator)
        return server

    def _isNewBatch(self, batch):
        new = self._batch != batch
        if new:
            self._dispose()
        return new

    def _needThread(self, user, merge):
        return merge and user.supportIMAP()

    def _getSender(self):
        return self._metadata.get('Sender')

    def _getSubject(self):
        subject = self._metadata.get('Subject')
        if self._merge:
            fields = self._getSubjectFields()
            try:
                subject = subject.format(**fields)
            except ValueError as e:
                pass
        return subject

    def _mergeMail(self, filter):
        descriptor = self._getFilteredDescriptor(filter)
        self._url.merge(descriptor)

    def _getDocumentTitle(self):
        return self._url.Title

    def _getAttachments(self, tag, separator):
        urls = []
        for url in self._urls:
            urls.append(tag % (url.Url, url.Name))
        return separator.join(urls)

    def _addAttachments(self, mail, filter):
        for url in self._urls:
            if self._merge and url.Merge:
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
        attachment.Data = self._transferable.getByUrl(url.Main)
        attachment.ReadableName = url.Name
        return attachment

    def _getDescriptors(self, merge, metadata):
        rowset = None
        descriptor = None
        if merge:
            name = metadata.get('DataSource')
            dbcontext, location = getDataBaseContext(self._ctx, self._source, name, self._getErrorMessage, 1531)
            datasource = dbcontext.getByName(name)
            connection = self._getDataSourceConnection(datasource)
            table = metadata.get('Table')
            if not connection.getTables().hasByName(table):
                msg = self._getErrorMessage(1534, name, table)
                raise MailSpoolerException(msg, self._source, ())
            rowset = self._getRowSet(connection, name, table)
            descriptor = {'DataSourceName': name,
                          'ActiveConnection': connection,
                          'Command': table,
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

    def _getRowSet(self, connection, name, table):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.ActiveConnection = connection
        rowset.DataSourceName = name
        rowset.CommandType = TABLE
        rowset.FetchSize = g_fetchsize
        rowset.Command = table
        return rowset

    def _getMailUrl(self, recipient, url, merge):
        uri = self._uf.parse(url)
        descriptor = self._getFilteredDescriptor(recipient.Filter, merge)
        return MailUrl(self._ctx, uri, merge, 'html', descriptor)

    def _getFilteredDescriptor(self, filter, merge=True):
        if merge:
            self._rowset.ApplyFilter = False
            self._rowset.Filter = filter
            self._rowset.ApplyFilter = True
            self._rowset.execute()
            self._descriptor['Cursor'] = self._rowset.createResultSet()
            descriptor = getPropertyValueSet(self._descriptor)
        else:
            descriptor = None
        return descriptor

    def _getMailAttachmentUrl(self, batch, job):
        urls = []
        for attachment in self._database.getAttachments(batch):
            url = self._uf.parse(attachment)
            merge, filter = self._getUrlMark(url)
            self._checkUrl(url.UriReference, job, 1521)
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
            msg = self._getErrorMessage(resource, job, url)
            raise MailSpoolerException(msg, self._source, ())

    def _getErrorMessage(self, code, *format):
        return self._logger.resolveString(code, *format)

