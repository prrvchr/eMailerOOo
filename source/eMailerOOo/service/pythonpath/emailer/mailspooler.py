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

from com.sun.star.ucb.ConnectionMode import OFFLINE

from com.sun.star.mail.MailServiceType import IMAP
from com.sun.star.mail.MailServiceType import SMTP

from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.uno import Exception as UnoException

from .oauth2 import getOAuth2Version

from .unotool import createService
from .unotool import executeDispatch
from .unotool import executeFrameDispatch
from .unotool import getConfiguration
from .unotool import getConnectionMode
from .unotool import getDesktop
from .unotool import getDocument
from .unotool import getExtensionVersion
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

from .mailertool import getDataSource
from .mailertool import getMail
from .mailertool import getMessageImage
from .mailertool import getUrlMimeType
from .mailertool import saveDocumentAs
from .mailertool import saveDocumentTmp
from .mailertool import saveTempDocument

from .logger import getLogger
from .logger import RollerHandler

from .dbconfig import g_version

from .configuration import g_dns
from .configuration import g_extension
from .configuration import g_identifier
from .configuration import g_spoolerlog
from .configuration import g_logo
from .configuration import g_logourl
from .configuration import g_fetchsize

from threading import Thread
from threading import Lock
from threading import Event
import traceback


class MailSpooler():
    def __init__(self, ctx):
        self._ctx = ctx
        self._lock = Lock()
        self._stop = Event()
        self._listeners = []
        logo = '%s/%s' % (g_extension, g_logo)
        self._logo = getResourceLocation(ctx, g_identifier, logo)
        self._logger = getLogger(ctx, g_spoolerlog)
        self._thread = Thread(target=self._run)
        self.DataSource = None

# Procedures called by TerminateListener and SpoolerService
    def stop(self):
        with self._lock:
            if self._thread.is_alive():
                self._stop.set()
                self._thread.join()

# Procedures called by SpoolerService
    def isStarted(self):
        with self._lock:
            return self._thread.is_alive()

    def start(self):
        with self._lock:
            if not self._thread.is_alive():
                self._thread = Thread(target=self._run)
                self._stop.clear()
                self._thread.start()
                self._started()

    def addListener(self, listener):
        self._listeners.append(listener)

    def removeListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

    def addJob(self, sender, subject, document, recipients, attachments):
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('addJob')
        if self._isInitialized():
            self.DataSource.waitForDataBase()
            return self.DataSource.insertJob(sender, subject, document, recipients, attachments)

    def addMergeJob(self, sender, subject, document, datasource, query, table, recipients, filters, attachments):
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('addMergeJob')
        if self._isInitialized():
            self.DataSource.waitForDataBase()
            return self.DataSource.insertMergeJob(sender, subject, document, datasource, query, table, recipients, filters, attachments)

    def removeJobs(self, jobids):
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('removeJobs')
        if self._isInitialized():
            self.DataSource.waitForDataBase()
            return self.DataSource.deleteJob(jobids)

    def getJobState(self, jobid):
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('getJobState')
        if self._isInitialized():
            self.DataSource.waitForDataBase()
            return self.DataSource.getJobState(jobid)

    def getJobIds(self, batchid):
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('getJobIds')
        if self._isInitialized():
            self.DataSource.waitForDataBase()
            return self.DataSource.getJobIds(batchid)

# Private methods
    def _isInitialized(self):
        return self.DataSource is not None

    def _initialize(self, method):
        self.DataSource = getDataSource(self._ctx, method, ' ', self._logError)

    def _logError(self, method, code, *args):
        self._logger.logprb(SEVERE, 'MailSpooler', method, code, *args)

    def _run(self):
        handler = RollerHandler(self._ctx, self._logger.Name)
        self._logger.addRollerHandler(handler)
        self._logger.logprb(INFO, 'MailSpooler', 'run()', 1001)
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('start')
        if self._isInitialized():
            # FIXME: Configuration has been checked we can continue
            if self._isOffLine():
                self._logger.logprb(INFO, 'MailSpooler', 'run()', 1002)
            else:
                self._logger.logprb(INFO, 'MailSpooler', 'run()', 1003)
                try:
                    self.DataSource.waitForDataBase()
                    connection = self.DataSource.DataBase.getConnection()
                    jobs, total = self.DataSource.DataBase.getSpoolerJobs(connection)
                    if total > 0:
                        configuration = getConfiguration(self._ctx, g_identifier, False)
                        timeout = configuration.getByName('ConnectTimeout')
                        self._logger.logprb(INFO, 'MailSpooler', 'run()', 1011, total)
                        count = self._send(connection, jobs, timeout)
                        self._logger.logprb(INFO, 'MailSpooler', 'run()', 1012, count, total)
                    else:
                        self._logger.logprb(INFO, 'MailSpooler', 'run()', 1013)
                    connection.close()
                except UnoException as e:
                    self._logger.logprb(SEVERE, 'MailSpooler', 'run()', 1014, e.Message)
                except Exception as e:
                    msg = "Error: %s" % traceback.format_exc()
                    self._logger.logprb(SEVERE, 'MailSpooler', 'run()', 1014, msg)
        self._logger.logprb(INFO, 'MailSpooler', 'run()', 1015)
        self._logger.removeRollerHandler(handler)
        self._stopped()

    def _started(self):
        event = uno.createUnoStruct('com.sun.star.lang.EventObject')
        for listener in self._listeners:
            listener.started(event)

    def _stopped(self):
        event = uno.createUnoStruct('com.sun.star.lang.EventObject')
        for listener in self._listeners:
            listener.stopped(event)

    def _send(self, connection, jobs, timeout):
        mailer = Mailer(self._ctx, self.DataSource.DataBase, self._logger)
        server = None
        count = 0
        for job in jobs:
            if self._stop.is_set():
                self._logger.logprb(INFO, 'MailSpooler', '_send()', 1024, job)
                break
            self._logger.logprb(INFO, 'MailSpooler', '_send()', 1021, job)
            recipient = self.DataSource.DataBase.getRecipient(connection, job)
            batch = recipient.BatchId
            if mailer.isNewBatch(batch):
                try:
                    metadata = self.DataSource.DataBase.getMailer(connection, batch, timeout)
                    if metadata is None:
                        self._logger.logprb(INFO, 'MailSpooler', '_send()', 1024, job)
                        continue
                    attachments = self.DataSource.DataBase.getAttachments(connection, batch)
                    mailer.setBatch(batch, metadata, attachments, job, recipient.Filter)
                    if self._stop.is_set():
                        self._logger.logprb(INFO, 'MailSpooler', '_send()', 1024, job)
                        break
                    if mailer.needThreadId():
                        self._createThreadId(connection, mailer, batch)
                    server = self._getServer(SMTP, mailer.getConfig(), server)
                except UnoException as e:
                    self.DataSource.DataBase.setJobState(connection, 2, job)
                    self._logger.logprb(INFO, 'MailSpooler', '_send()', 1022, e.Message)
                    continue
            elif mailer.Merge:
                mailer.merge(recipient.Filter)
            if self._stop.is_set():
                self._logger.logprb(INFO, 'MailSpooler', '_send()', 1024, job)
                break
            mail = self._getMail(mailer, recipient)
            mailer.addAttachments(mail, recipient.Filter)
            if self._stop.is_set():
                self._logger.logprb(INFO, 'MailSpooler', '_send()', 1024, job)
                break
            server.sendMailMessage(mail)
            self.DataSource.DataBase.updateRecipient(connection, 1, mail.MessageId, job)
            self._logger.logprb(INFO, 'MailSpooler', '_send()', 1023, job)
            count += 1
        self._dispose(server, mailer)
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
        title = self._logger.resolveString(1031, g_extension, batch, mailer.Query)
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

    def _isOffLine(self):
        mode = getConnectionMode(self._ctx, *g_dns)
        return mode == OFFLINE


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
        return self._metadata.get('Merge')
    @property
    def Sender(self):
        return self._metadata.get('Sender')
    @property
    def Subject(self):
        return self._metadata.get('Subject')
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
        self._metadata = metadata
        self._checkUrl(self.Document, job, 161)
        self._rowset, self._descriptor = self._getDescriptors()
        self._urls, self._url = self._getUrls(attachments, job, filter)

    def needThreadId(self):
        return self.ThreadId is None and self.Merge and self._hasImapConfig()

    def setThreadId(self, connection, batchid, thread):
        self._database.updateMailer(connection, batchid, thread)
        self._metadata['ThreadId'] = thread

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
        return all((self._metadata.get('IMAPServerName'),
                    self._metadata.get('IMAPPort'),
                    self._metadata.get('IMAPLogin'),
                    self._metadata.get('IMAPConnectionType'),
                    self._metadata.get('IMAPAuthenticationType')))

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

