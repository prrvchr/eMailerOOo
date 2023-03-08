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

from com.sun.star.mail.MailServiceType import IMAP
from com.sun.star.mail.MailServiceType import SMTP

from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.uno import Exception as UnoException

from .unotool import createService
from .unotool import executeDispatch
from .unotool import executeFrameDispatch
from .unotool import getConfiguration
from .unotool import getConnectionMode
from .unotool import getDesktop
from .unotool import getDocument
from .unotool import getFileSequence
from .unotool import getInteractionHandler
from .unotool import getPropertyValueSet
from .unotool import getResourceLocation
from .unotool import getSimpleFile
from .unotool import getTempFile
from .unotool import getTypeDetection
from .unotool import getUriFactory

from .mailerlib import Authenticator
from .mailerlib import CurrentContext
from .mailerlib import MailTransferable

from .listener import TerminateListener

from .mailertool import getMail
from .mailertool import getMessageImage
from .mailertool import getUrlMimeType
from .mailertool import saveDocumentAs
from .mailertool import saveDocumentTmp
from .mailertool import saveTempDocument

from .logger import getLogger

from .configuration import g_dns
from .configuration import g_extension
from .configuration import g_identifier
from .configuration import g_spoolerlog
from .configuration import g_logo
from .configuration import g_logourl
from .configuration import g_fetchsize

from threading import Thread
from threading import Condition
from threading import Event
import traceback
import time


class MailSpooler(Thread):
    def __init__(self, ctx, database, listeners):
        Thread.__init__(self)
        self._ctx = ctx
        self._database = database
        self._lock = Condition()
        self._disposed = Event()
        self._disposed.clear()
        self._listeners = listeners
        logo = '%s/%s' % (g_extension, g_logo)
        self._logo = getResourceLocation(ctx, g_identifier, logo)
        #self._listener = TerminateListener(self)
        #self._terminated = False
        #getDesktop(ctx).addTerminateListener(self._listener)
        configuration = getConfiguration(ctx, g_identifier, False)
        self._timeout = configuration.getByName('ConnectTimeout')
        self._logger = getLogger(ctx, g_spoolerlog)

    # Thread
    def run(self):
        try:
            print("MailSpooler._run() **********************************")
            connection = self._database.getConnection()
            print("MailSpooler._run() 1")
            jobs, total = self._database.getSpoolerJobs(connection)
            print("MailSpooler._run() 2")
            if total > 0:
                print("MailSpooler._run() 3")
                self._logger.logprb(INFO, 'MailSpooler', 'run()', 1011, total)
                count = self._send(connection, jobs)
                print("MailSpooler._run() 4")
                self._logger.logprb(INFO, 'MailSpooler', 'run()', 1012, count, total)
            else:
                print("MailSpooler._run() 5")
                self._logger.logprb(INFO, 'MailSpooler', 'run()', 1013)
            print("MailSpooler._run() 6")
            connection.close()
        except UnoException as e:
            self._logger.logprb(INFO, 'MailSpooler', 'run()', 1014, e.Message)
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)
        print("MailSpooler._run() 7")
        self._logger.logprb(INFO, 'MailSpooler', 'run()', 1015)
        print("MailSpooler._run() 8")
        self._stopped()
        #if not self._terminated:
        #    getDesktop(self._ctx).removeTerminateListener(self._listener)
        print("MailSpooler._run() 9")

# Procedures called by SpoolerService
    def clearDisposed(self):
        self._disposed.clear()

    def setDisposed(self):
        self._disposed.set()

    def addListener(self, listener):
        with self._lock:
            print("MailSpooler.addListener()")
            self._listeners.append(listener)

    def removeListener(self, listener):
        with self._lock:
            if listener in self._listeners:
                print("MailSpooler.removeListener()")
                self._listeners.remove(listener)

    def started(self):
        with self._lock:
            event = uno.createUnoStruct('com.sun.star.lang.EventObject')
            for listener in self._listeners:
                listener.started(event)

# Procedures called by TerminateListener
    def stop(self):
        print("MailSpooler.stop() 1")
        if self.is_alive():
            print("MailSpooler.stop() 2")
            self._terminated = True
            self._disposed.set()
            print("MailSpooler.stop() 3")
            self.join()
        print("MailSpooler.stop() 4")

# Private methods
    def _stopped(self):
        with self._lock:
            event = uno.createUnoStruct('com.sun.star.lang.EventObject')
            for listener in self._listeners:
                listener.stopped(event)

    def _send(self, connection, jobs):
        print("MailSpooler._send() 1 begin ****************************************")
        mailer = Mailer(self._ctx, self._database, self._logger)
        server = None
        count = 0
        for job in jobs:
            if self._disposed.is_set():
                self._logger.logprb(INFO, 'MailSpooler', '_send()', 1024, job)
                break
            self._logger.logprb(INFO, 'MailSpooler', '_send()', 1021, job)
            recipient = self._database.getRecipient(connection, job)
            batch = recipient.BatchId
            if mailer.isNewBatch(batch):
                try:
                    print("MailSpooler._send() 2")
                    metadata = self._database.getMailer(connection, batch, self._timeout)
                    if metadata is None:
                        self._logger.logprb(INFO, 'MailSpooler', '_send()', 1024, job)
                        continue
                    print("MailSpooler._send() 3")
                    attachments = self._database.getAttachments(connection, batch)
                    print("MailSpooler._send() 4")
                    mailer.setBatch(batch, metadata, attachments, job, recipient.Filter)
                    print("MailSpooler._send() 5")
                    if self._disposed.is_set():
                        self._logger.logprb(INFO, 'MailSpooler', '_send()', 1024, job)
                        break
                    if mailer.needThreadId():
                        print("MailSpooler._send() 6")
                        self._createThreadId(connection, mailer, batch)
                    server = self._getServer(SMTP, mailer.getConfig(), server)
                    print("MailSpooler._send() 7")
                except UnoException as e:
                    print("MailSpooler._send() 8 break %s" % e.Message)
                    self._database.setJobState(connection, 2, job)
                    self._logger.logprb(INFO, 'MailSpooler', '_send()', 1022, e.Message)
                    print("MailSpooler._send() 9 break %s" % msg)
                    continue
            elif mailer.Merge:
                mailer.merge(recipient.Filter)
            mail = self._getMail(mailer, recipient)
            mailer.addAttachments(mail, recipient.Filter)
            print("MailSpooler._send() 10 Filter: %s" % recipient.Filter)
            if self._disposed.is_set():
                print("MailSpooler._send() 11 break")
                self._logger.logprb(INFO, 'MailSpooler', '_send()', 1024, job)
                break
            server.sendMailMessage(mail)
            self._database.updateRecipient(connection, 1, mail.MessageId, job)
            self._logger.logprb(INFO, 'MailSpooler', '_send()', 1023, job)
            count += 1
            print("MailSpooler._send() 12")
        self._dispose(server, mailer)
        print("MailSpooler._send() 13 end.........................")
        return count

    def _createThreadId(self, connection, mailer, batch):
        server = self._getServer(IMAP, mailer.getConfig())
        folder = server.getSentFolder()
        if server.hasFolder(folder):
            message = self._getThreadMessage(mailer, batch)
            body = MailTransferable(self._ctx, message, True)
            mail = getMail(self._ctx, mailer.Sender, mailer.Sender, mailer.Subject, body)
            server.uploadMessage(folder, mail)
            mailer.setThreadId(connection, batch, mail.MessageId)
        server.disconnect()

    def _getThreadMessage(self, mailer, batch):
        title = self._logger.resolveString(1031, batch, mailer.Query)
        subject = self._logger.resolveString(1032)
        document = self._logger.resolveString(1033)
        files = self._logger.resolveString(1034)
        if mailer.hasAttachments():
            tag = '<a href="%s">%s</a>'
            separator = '</li><li>'
            attachments = '<ol><li>%s</li></ol>' % mailer.getAttachments(tag, separator)
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
''' % (g_extension, logo, g_logourl, title, subject, mailer.Subject,
        document, mailer.Document, mailer.getDocumentTitle(),
        files, attachments)

    def _dispose(self, server, mailer):
        if server is not None:
            server.disconnect()
        mailer.dispose()

    def _getMail(self, mailer, recipient):
        body = MailTransferable(self._ctx, mailer.getBodyUrl(), True, True)
        mail = getMail(self._ctx, mailer.Sender, recipient.Recipient, mailer.Subject, body)
        return mail

    def _getServer(self, mailtype, config, server=None):
        if server is not None:
            server.disconnect()
        server = self._getMailServer(mailtype, config)
        return server

    def _checkUrl(self, sf, url, job, resource):
        if not sf.exists(url):
            raise _getUnoException(self._logger, self, resource, job, url)

    def _getMailServer(self, mailtype, config):
        context = CurrentContext(mailtype.value, config)
        authenticator = Authenticator(mailtype.value, config)
        service = 'com.sun.star.mail.MailServiceProvider2'
        try:
            host = context.getValueByName('ServerName')
            server = createService(self._ctx, service).create(mailtype, host)
            server.connect(context, authenticator)
        except UnoException as e:
            mailserver = '%s:%s' % (server['ServerName'], server['Port'])
            raise _getUnoException(self._logger, self, 1041, job, sender.Sender, mailserver, e.Message)
        return server


class Mailer(unohelper.Base):
    def __init__(self, ctx, database, logger):
        self._ctx = ctx
        self._database = database
        self._logger = logger
        self._sf = getSimpleFile(ctx)
        self._uf = getUriFactory(ctx)
        self._batch = None
        self._descriptor = None
        self._url = None
        self._urls = ()
        self._thread = None
        self._rowset = None

    @property
    def Merge(self):
        return self._metadata.getValue('Merge')
    @property
    def Sender(self):
        return self._metadata.getValue('Sender')
    @property
    def Subject(self):
        return self._metadata.getValue('Subject')
    @property
    def Document(self):
        return self._metadata.getValue('Document')
    @property
    def DataSource(self):
        return self._metadata.getValue('DataSource')
    @property
    def Table(self):
        return self._metadata.getValue('Table')
    @property
    def Query(self):
        return self._metadata.getValue('Query')
    @property
    def ThreadId(self):
        return self._metadata.getValue('ThreadId')

    def isNewBatch(self, batch):
        new = self._batch != batch
        if new and self._batch is not None:
            self._dispose()
        return new

    def setBatch(self, batch, metadata, attachments, job, filter):
        self._batch = batch
        self._metadata = metadata
        self._checkUrl(self.Document, job, 161)
        self._rowset, self._descriptor = self._getDescriptors()
        self._urls, self._url = self._getUrls(attachments, job, filter)

    def needThreadId(self):
        return self.ThreadId is None and self.Merge and self._hasImapConfig()

    def setThreadId(self, connection, batchid, thread):
        self._database.updateMailer(connection, batchid, thread)
        self._metadata.setValue('ThreadId', thread)

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

    def getConfig(self):
        return self._metadata

# Private Procedures call
    def _dispose(self):
        if self._descriptor is not None:
            self._descriptor['ActiveConnection'].close()
        self._url.dispose()
        for url in self._urls:
            url.dispose()

    def _hasImapConfig(self):
        return all((self._metadata.getValue('IMAPServerName'),
                    self._metadata.getValue('IMAPPort'),
                    self._metadata.getValue('IMAPLogin'),
                    self._metadata.getValue('IMAPConnectionType'),
                    self._metadata.getValue('IMAPAuthenticationType')))

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
            raise _getUnoException(self._logger, self, resource, job, url)


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


def _getUnoException(logger, source, resource, *args):
    e = UnoException()
    e.Message = logger.resolveString(resource, *args)
    e.Context = source
    return e
