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

from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.mail.MailServiceType import IMAP

from com.sun.star.mail import MailSpoolerException

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.sdb.CommandType import TABLE

from .mailurl import MailUrl

from ..merger import Merger

from ..export import PdfExport

from ....transferable import Transferable

from ....unotool import createService
from ....unotool import getInteractionHandler
from ....unotool import getResourceLocation
from ....unotool import getSimpleFile
from ....unotool import getUriFactory
from ....unotool import hasInterface

from ....dbtool import getRowDict

from ....helper import getDataBaseContext
from ....helper import getMailMessage
from ....helper import getMailService
from ....helper import getMailUser
from ....helper import getMessageImage

from ....configuration import g_extension
from ....configuration import g_identifier
from ....configuration import g_logourl
from ....configuration import g_logo

import traceback


class Mailer():
    def __init__(self, ctx, source, database, logger, send=False, stop=None):
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
        self._url = None
        self._urls = ()
        self._merger = None
        self._rowset = None
        self._stop = stop
        self._export = PdfExport(ctx)
        self._threadid = None
        self._reply = False
        self._replyto = ''
        logo = '%s/%s' % (g_extension, g_logo)
        self._logo = getResourceLocation(ctx, g_identifier, logo)
        self._logger.logprb(INFO, 'Mailer', '__init__', 1501)

    @property
    def _merge(self):
        return self._metadata.get('Merge')

    def dispose(self):
        self._dispose()
        self._logger.logprb(INFO, 'Mailer', 'dispose', 1561)

    def getMail(self, job):
        batch, mail = None, None
        recipient = self._database.getRecipient(job)
        batch = recipient.BatchId
        if self._isCancelled():
            return batch, mail
        if self._isNewBatch(batch):
            self._initMailer(job, recipient)
        elif self._merge:
            self._mergeUrl(job, self._url, recipient.Filter, 1517)
        if self._isCancelled():
            return batch, mail
        body = self._transferable.getByUrl(self._url.Main)
        if self._isCancelled():
            return batch, mail
        mail = getMailMessage(self._ctx, self._getSender(), recipient.Recipient, self._getSubject(recipient.Filter), body)
        if not self._isCancelled():
            if self._hasThread():
                mail.ThreadId = self._threadid
            if self._reply:
                mail.ReplyToAddress = self._replyto
            if not self._isCancelled():
                for url in self._urls:
                    if self._isCancelled():
                        break
                    if url.Merge:
                        self._mergeUrl(job, url, recipient.Filter, 1525)
                    mail.addAttachment(self._getAttachment(url))
        return batch, mail

    def sendMail(self, mail):
        self._server.sendMailMessage(mail)

# Private Procedures call
    def _isCancelled(self):
        return self._stop and self._stop.is_set()

    def _dispose(self):
        if self._server is not None:
            self._server.disconnect()
        if self._url is not None:
            self._url.dispose(self._sf)
        for url in self._urls:
            url.dispose(self._sf)
        if self._merger:
            self._merger.dispose()
            self._merger = None
        self._rowset = None

    def _initMailer(self, job, recipient):
        threadid = None
        metadata = self._database.getMailer(recipient.BatchId)
        url, merge, sender, connection, datasource, table = self._getMetaData(job, metadata)

        if self._isCancelled():
            return
        self._checkUrl(job, url, 1513)
        self._url = self._getMailUrl(job, recipient, url, merge, connection, datasource, table)
        self._urls = self._getMailAttachmentUrl(job, recipient.BatchId, connection, datasource, table)

        if self._isCancelled():
            return
        user = self._getMailUser(job, sender)
        self._reply = user.useReplyTo()
        self._replyto = user.getReplyToAddress()

        if self._isCancelled():
            return
        if self._send:
            self._server = self._getMailServer(user, SMTP)
            if self._needThread(user, merge):
                threadid = self._createThread(recipient.BatchId, user, url, sender, metadata)
        if self._isCancelled():
            return
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
            # XXX: Some providers use their own reference (ie: Microsoft with their ConversationId)
            # XXX: in order to group the mails and not the MessageId of the first mail.
            # XXX: If so, the ThreadId is not empty.
            threadid = mail.ForeignId if mail.ForeignId else mail.MessageId
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
       label, url, self._url.Title, files, attachments)

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

    def _getSubject(self, filter):
        subject = self._metadata.get('Subject')
        if self._merge and filter:
            fields = self._getSubjectFields(filter)
            try:
                subject = subject.format(**fields)
            except ValueError as e:
                pass
        return subject

    def _getSubjectFields(self, filter):
        fields = {}
        result = self._getFilteredResultSet(filter)
        if result.next():
            fields = getRowDict(result, '')
        result.close()
        return fields

    def _mergeUrl(self, job, url, filter, resource):
        try:
            self._merger.merge(self._sf, url, filter)
        except Exception as e:
            self._raiseSpoolerException(resource, job, url.Document.Title, e)

    def _getAttachments(self, tag, separator):
        urls = []
        for url in self._urls:
            urls.append(tag % (url.Url, url.Name))
        return separator.join(urls)

    def _addAttachments(self, job, mail, filter):
        for url in self._urls:
            if self._isCancelled():
                break
            if url.Merge:
                self._mergeUrl(job, url, filter, 1525)
            mail.addAttachment(self._getAttachment(url))

    def _hasAttachments(self):
        return len(self._urls) > 0

    def _hasThread(self):
        return self._threadid is not None

    def _getAttachment(self, url):
        attachment = uno.createUnoStruct('com.sun.star.mail.MailAttachment')
        attachment.Data = self._transferable.getByUrl(url.Main)
        attachment.ReadableName = url.Name
        return attachment

    def _getMetaData(self, job, metadata):
        url = metadata.get('Document')
        merge = metadata.get('Merge')
        sender = metadata.get('Sender')
        merger = rowset = connection = datasource = table = None
        if merge:
            print("Mailer._getMetaData() 1")
            datasource = metadata.get('DataSource')
            connection = self._getDataSourceConnection(datasource)
            if not self._checkConnectionInterface(connection):
                self._raiseSpoolerException(1511, job, datasource)
            print("Mailer._getMetaData() 2")
            table = metadata.get('Table')
            if not self._checkConnectionTable(connection, table):
                self._raiseSpoolerException(1512, job, datasource, table)
            print("Mailer._getMetaData() 3")
            merger = Merger(self._ctx, connection, datasource, table, self._export)
            rowset = self._getRowSet(connection, datasource, table)
        self._merger = merger
        self._rowset = rowset
        print("Mailer._getMetaData() 4")
        return url, merge, sender, connection, datasource, table

    def _getRowSet(self, connection, datasource, table):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.ActiveConnection = connection
        rowset.DataSourceName = datasource
        rowset.CommandType = TABLE
        rowset.Command = table
        return rowset

    def _getFilteredResultSet(self, filter):
        self._rowset.ApplyFilter = False
        self._rowset.Filter = filter
        self._rowset.ApplyFilter = True
        self._rowset.execute()
        return self._rowset.createResultSet()

    def _getDataSourceConnection(self, name):
        print("Mailer._getDataSourceConnection() 1")
        dbcontext, location = getDataBaseContext(self._ctx, self._source, name, self._getErrorMessage, 1531)
        datasource = dbcontext.getByName(name)
        if datasource.IsPasswordRequired:
            print("Mailer._getDataSourceConnection() 2")
            handler = getInteractionHandler(self._ctx)
            connection = datasource.getIsolatedConnectionWithCompletion(handler)
            print("Mailer._getDataSourceConnection() 3")
        else:
            print("Mailer._getDataSourceConnection() 4")
            connection = datasource.getIsolatedConnection('', '')
        return connection

    def _getMailUrl(self, job, recipient, url, merge, connection, datasource, table):
        uri = self._uf.parse(url)
        mail = self._getMail(job, uri, merge, 'html', 1514)
        print("Mailer._getMailUrl() 1")
        if mail.isDocumentInvalid():
            self._raiseSpoolerException(1515, job, url)
        print("Mailer._getMailUrl() 2 merge: %s" % merge)
        if merge:
            if self._merger.hasInvalidFields(mail.Document, connection, datasource, table):
                self._raiseSpoolerException(1516, job, *self._merger.getInvalidFields())
            print("Mailer._getMailUrl() 3 document merge ok")
            self._mergeUrl(job, mail, recipient.Filter, 1517)
        else:
            mail.export()
        print("Mailer._getMailUrl() 5")
        return mail

    def _getMailAttachmentUrl(self, job, batch, connection, datasource, table):
        urls = []
        for attachment in self._database.getAttachments(batch):
            if self._isCancelled():
                break
            url = self._uf.parse(attachment)
            merge, filter = self._getUrlMark(url)
            self._checkUrl(job, url.UriReference, 1521)
            mail = self._getMail(job, url, merge, filter, 1522)
            print("Mailer._getMailAttachmentUrl() 1")
            if mail.isDocumentInvalid():
                self._raiseSpoolerException(1523, job, url)
            print("Mailer._getMailAttachmentUrl() 2")
            if merge:
                if self._merger is None:
                    self._merger = Merger(self._ctx, connection, datasource, table, self._export)
                    self._rowset = self._getRowSet(connection, datasource, table)
                if self._merger.hasInvalidFields(mail.Document, connection, datasource, table):
                    self._raiseSpoolerException(1524, job, *self._merger.getInvalidFields())
                print("Mailer._getMailAttachmentUrl() 3 attachment merge ok")
            elif filter:
                mail.export(self._export)
            urls.append(mail)
        return urls

    def _getMail(self, job, url, merge, filter, resource):
        try:
            return MailUrl(self._ctx, url, merge, filter)
        except Exception as e:
            self._raiseSpoolerException(resource, job, url.UriReference, e)

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

    def _checkUrl(self, job, url, resource):
        if not self._sf.exists(url):
            self._raiseSpoolerException(resource, job, url)

    def _checkConnectionInterface(self, connection):
        return hasInterface(connection, 'com.sun.star.sdbcx.XTablesSupplier')

    def _checkConnectionTable(self, connection, table):
        return connection.getTables().hasByName(table)

    def _checkDocument(self, mail, job, url):
        if mail.isDocumentNotFound(self._sf):
            self._raiseSpoolerException(1512, job, url)

    def _raiseSpoolerException(self, resource, *args):
        msg = self._getErrorMessage(resource, *args)
        raise MailSpoolerException(msg, self._source, ())

    def _getErrorMessage(self, code, *args):
        return self._logger.resolveString(code, *args)

