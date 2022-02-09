#!
# -*- coding: utf_8 -*-

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

from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.uno import Exception as UnoException

from .unotool import createService
from .unotool import executeDispatch
from .unotool import executeFrameDispatch
from .unotool import getConfiguration
from .unotool import getConnectionMode
from .unotool import getFileSequence
from .unotool import getInteractionHandler
from .unotool import getPropertyValueSet
from .unotool import getSimpleFile
from .unotool import getTempFile
from .unotool import getTypeDetection
from .unotool import getUriFactory

from smtpmailer import MailTransferable
from smtpmailer import TerminateListener

from smtpmailer import getMail
from smtpmailer import getDesktop
from smtpmailer import getDocument
from smtpmailer import getUrl
from smtpmailer import getUrlMimeType
from smtpmailer import saveDocumentAs
from smtpmailer import saveDocumentTmp
from smtpmailer import saveTempDocument
from smtpmailer import Authenticator
from smtpmailer import CurrentContext

from .logger import Pool

from .configuration import g_dns
from .configuration import g_identifier
from .configuration import g_fetchsize

#from multiprocessing.context import ForkProcess as Process
from threading import Thread as Process
from threading import Condition
from threading import Event
import traceback
import time


class MailSpooler(Process):
    def __init__(self, ctx, database, logger, listeners):
        Process.__init__(self)
        self._ctx = ctx
        self._database = database
        self._lock = Condition()
        self._disposed = Event()
        self._disposed.clear()
        self._listeners = listeners
        #self._listener = TerminateListener(self)
        #self._terminated = False
        #getDesktop(ctx).addTerminateListener(self._listener)
        configuration = getConfiguration(ctx, g_identifier, False)
        self._timeout = configuration.getByName('ConnectTimeout')
        self._logger = logger

# Process
    def run(self):
        try:
            print("MailSpooler._run() **********************************")
            connection = self._database.getConnection()
            print("MailSpooler._run() 1")
            jobs, total = self._database.getSpoolerJobs(connection)
            print("MailSpooler._run() 2")
            if total > 0:
                print("MailSpooler._run() 3")
                self._logger.logResource(INFO, 111, total)
                count = self._send(connection, jobs)
                format = (count, total)
                print("MailSpooler._run() 4")
                self._logger.logResource(INFO, 112, format)
            else:
                print("MailSpooler._run() 5")
                self._logger.logResource(INFO, 113)
            print("MailSpooler._run() 6")
            connection.close()
        except UnoException as e:
            msg = self._logger.getMessage(114, e.Message)
            print(msg)
            self._logger.logMessage(INFO, msg)
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)
        print("MailSpooler._run() 7")
        self._logger.logResource(INFO, 115)
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
        print("MailSpooler._send()1 begin ****************************************")
        mailer = Mailer(self._ctx, self._database, self._logger)
        server = None
        count = 0
        for job in jobs:
            if self._disposed.is_set():
                self._logger.logResource(INFO, 124, job)
                break
            self._logger.logResource(INFO, 121, job)
            recipient = self._database.getRecipient(connection, job)
            batch = recipient.BatchId
            if mailer.isNewBatch(batch):
                try:
                    metadata = self._database.getMailer(connection, batch, self._timeout)
                    if metadata is None:
                        # TODO: need to log this user has no IspDB config
                        self._logger.logResource(INFO, 124, job)
                        continue
                    attachments = self._database.getAttachments(connection, batch)
                    mailer.setBatch(batch, metadata, attachments, job, recipient.Index)
                    if self._disposed.is_set():
                        self._logger.logResource(INFO, 124, job)
                        break
                    if mailer.hasThread():
                        self._createThread(mailer, batch)
                    server = self._getServer(SMTP, mailer.getConfig(), server)
                    print("MailSpooler._send() 4")
                except UnoException as e:
                    print("MailSpooler._send() 5 break %s" % e.Message)
                    self._database.setJobState(connection, 2, job)
                    msg = self._logger.getMessage(122, e.Message)
                    self._logger.logMessage(INFO, msg)
                    print("MailSpooler._send() 6 break %s" % msg)
                    continue
            elif mailer.Merge:
                mailer.merge(recipient.Index)
            mail = self._getMail(mailer, recipient)
            mailer.addAttachments(mail, recipient.Index)
            if self._disposed.is_set():
                print("MailSpooler._send() 8 break")
                self._logger.logResource(INFO, 124, job)
                break
            server.sendMailMessage(mail)
            self._database.setJobState(connection, 1, job)
            self._logger.logResource(INFO, 123, job)
            count += 1
            print("MailSpooler._send() 9")
        self._dispose(server, mailer)
        print("MailSpooler._send() 11 end.........................")
        return count

    def _createThread(self, mailer, batch):
        print("MailSpooler._createThread() 1 *****************************************************")
        server = self._getServer(IMAP, mailer.getConfig())
        folder = server.findSentFolder()
        if server.hasFolder(folder):
            subject = self._getThreadSubject(mailer, batch)
            message = self._getThreadMessage(mailer, subject)
            body = MailTransferable(self._ctx, message, False)
            mail = getMail(self._ctx, mailer.Sender, mailer.Sender, subject, body)
            server.uploadMessage(folder, mail)
            mailer.setThread(mail.MessageId)
        print("MailSpooler._createThread() 2 *****************************************************")
        server.disconnect()
        print("MailSpooler._createThread() 3 *****************************************************")

    def _getThreadSubject(self, mailer, batch):
        format = (batch, mailer.Query)
        subject = self._logger.getMessage(131, format)
        return subject

    def _getThreadMessage(self, mailer, title):
        subject = self._logger.getMessage(141)
        attachments = self._logger.getMessage(142)
        return '''\
%s

%s: %s

Document: %s

%s:
%s

''' % (title, subject, mailer.Subject, mailer.Document, attachments, mailer.getUrls())

    def _dispose(self, server, mailer):
        if server is not None:
            server.disconnect()
        mailer.dispose()

    def _getMail(self, mailer, recipient):
        body = MailTransferable(self._ctx, mailer.getBodyUrl(), True)
        mail = getMail(self._ctx, mailer.Sender, recipient.Recipient, mailer.Subject, body)
        return mail

    def _checkSender(self, sf, sender, job):
        if sender is None:
            raise _getUnoException(self._logger, self, 131, job)
        self._checkUrl(sf, sender.Document, job, 132)

    def _getServer(self, mailtype, config, server=None):
        if server is not None:
            server.disconnect()
        server = self._getMailServer(mailtype, config)
        return server

    def _checkUrl(self, sf, url, job, resource):
        if not sf.exists(url):
            format = (job, url)
            raise _getUnoException(self._logger, self, resource, format)

    def _getMailServer(self, mailtype, config):
        context = CurrentContext(mailtype.value, config)
        authenticator = Authenticator(mailtype.value, config)
        service = 'com.sun.star.mail.MailServiceProvider2'
        try:
            server = createService(self._ctx, service).create(mailtype)
            server.connect(context, authenticator)
        except UnoException as e:
            mailserver = '%s:%s' % (server['ServerName'], server['Port'])
            format = (job, sender.Sender, mailserver, e.Message)
            raise _getUnoException(self._logger, self, 151, format)
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
    def Identifier(self):
        return self._metadata.getValue('Identifier')
    @property
    def Bookmark(self):
        return self._metadata.getValue('Bookmark')

    def isNewBatch(self, batch):
        new = self._batch != batch
        if new and self._batch is not None:
            self._dispose()
        return new

    def setBatch(self, batch, metadata, attachments, job, index):
        self._batch = batch
        self._metadata = metadata
        self._checkUrl(self.Document, job, 161)
        self._descriptor = self._getDescriptor()
        self._urls, self._url = self._getUrls(attachments, job, index)

    def hasThread(self):
        self._thread = None
        return self.Merge and self._hasImapConfig()

    def setThread(self, thread):
        self._thread = thread

    def _hasImapConfig(self):
        return all((self._metadata.getValue('IMAPServerName'),
                    self._metadata.getValue('IMAPPort'),
                    self._metadata.getValue('IMAPLogin'),
                    self._metadata.getValue('IMAPConnectionType'),
                    self._metadata.getValue('IMAPAuthenticationType')))

    def dispose(self):
        if self._batch is not None:
            self._dispose()

    def _dispose(self):
        self._descriptor['ActiveConnection'].close()
        self._url.dispose()
        for url in self._urls:
            url.dispose()

    def merge(self, index):
        descriptor = self._getUrlDescriptor(index)
        self._url.merge(descriptor)

    def getBodyUrl(self):
        return self._url.Main

    def getUrls(self):
        urls = []
        for url in self._urls:
            urls.append(url.Url)
        return '\n'.join(urls)

    def addAttachments(self, mail, index):
        if self._hasThread():
            mail.ThreadId = self._thread
        for url in self._urls:
            print("Mailer.addAttachments() %s - %s" % (url.Name, url.Main))
            if self.Merge and url.Merge:
                descriptor = self._getUrlDescriptor(index)
                url.merge(descriptor)
            mail.addAttachment(self._getAttachment(url))

    def _hasThread(self):
        return self._thread is not None

    def _getAttachment(self, url):
        attachment = uno.createUnoStruct('com.sun.star.mail.MailAttachment')
        attachment.Data = MailTransferable(self._ctx, url.Main)
        attachment.ReadableName = url.Name
        return attachment

    def getConfig(self):
        return self._metadata

    def _getDescriptor(self):
        service = 'com.sun.star.sdb.DatabaseContext'
        datasource = createService(self._ctx, service).getByName(self.DataSource)
        connection = self._getDataSourceConnection(datasource)
        rowset = self._getRowSet(connection)
        rowset.execute()
        descriptor = {'ActiveConnection': connection,
                      'DataSourceName': self.DataSource,
                      'Command': self.Table,
                      'CommandType': TABLE,
                      'BookmarkSelection': False,
                      'Cursor': rowset.createResultSet()}
        return descriptor

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
        rowset.CommandType = TABLE
        rowset.FetchSize = g_fetchsize
        rowset.Command = self.Table
        return rowset

    def _getUrls(self, attachments, job, index):
        descriptor = self._getUrlDescriptor(index) if self.Merge else None
        uri = self._uf.parse(self.Document)
        url = MailUrl(self._ctx, uri, self.Merge, False, descriptor)
        urls = self._getMailUrl(attachments, job)
        return urls, url

    def _getUrlDescriptor(self, index):
        connection = self._descriptor['ActiveConnection']
        format = (self.Bookmark, self.Table, self.Identifier)
        bookmark = self._database.getBookmark(connection, format, index)
        self._descriptor['Selection'] = (bookmark, )
        descriptor = getPropertyValueSet(self._descriptor)
        return descriptor

    def _getMailUrl(self, attachments, job):
        urls = []
        for attachment in attachments:
            url = self._uf.parse(attachment)
            merge, pdf = self._getUrlMark(url)
            self._checkUrl(url.UriReference, job, 171)
            urls.append(MailUrl(self._ctx, url, merge, pdf))
        return urls

    def _getUrlMark(self, url):
        merge = pdf = False
        if url.hasFragment():
            fragment = url.getFragment()
            if fragment.startswith('merge'):
                merge = True
            if fragment.endswith('pdf'):
                pdf = True
            url.clearFragment()
        return merge, pdf

    def _checkUrl(self, url, job, resource):
        if not self._sf.exists(url):
            format = (job, url)
            raise _getUnoException(self._logger, self, resource, format)


class MailUrl(unohelper.Base):
    def __init__(self, ctx, url, merge, pdf, descriptor=None):
        self._ctx = ctx
        self._merge = merge
        self._html = descriptor is not None
        self._pdf = pdf
        self._url = url.UriReference
        name = url.getPathSegment(url.PathSegmentCount -1)
        self._name = self._title = name
        if self._isTemp():
            self._temp = getTempFile(ctx).Uri
            self._document = getDocument(ctx, self._url)
            if not self._merge:
                self._title = self._saveTempDocument()
            elif self._html:
                self.merge(descriptor)
        else:
            self._temp = self._document = None

    @property
    def Merge(self):
        return self._merge
    @property
    def Name(self):
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
        self._title = self._saveTempDocument()
        print("MailUrl.merge() %s - %s" % (self.Name, self.Main))

    def dispose(self):
        if self._isTemp():
            self._document.close(True)
            self._temp = None

# Private Procedures call
    def _isTemp(self):
        return any((self._merge, self._html, self._pdf))

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
        filter = None
        if self._html:
            filter = 'html'
        elif self._pdf:
            filter = 'pdf'
        return saveTempDocument(self._document, self._temp, self._name, filter)


def _getUnoException(logger, source, resource, format=None):
    e = UnoException()
    e.Message = logger.getMessage(resource, format)
    e.Context = source
    return e
