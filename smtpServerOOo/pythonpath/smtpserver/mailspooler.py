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

from com.sun.star.mail.MailServiceType import SMTP

from com.sun.star.ucb.ConnectionMode import OFFLINE

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
from .unotool import getDateTime

from smtpserver import MailTransferable

from smtpserver import getMail
from smtpserver import getDocument
from smtpserver import getUrl
from smtpserver import getUrlMimeType
from smtpserver import saveDocumentAs
from smtpserver import saveDocumentTmp
from smtpserver import saveTempDocument
from smtpserver import Authenticator
from smtpserver import CurrentContext

from .logger import Pool

from .configuration import g_dns
from .configuration import g_identifier
from .configuration import g_fetchsize

from multiprocessing import Process, Queue, cpu_count
from threading import Thread
from threading import Event
import traceback
import time


class MailSpooler(unohelper.Base):
    def __init__(self, ctx, database, event):
        self._ctx = ctx
        self._database = database
        self._disposed = Event()
        self._server = None
        self._thread = None
        self._descriptor = {}
        self._listeners = []
        self._event = event
        configuration = getConfiguration(ctx, g_identifier, False)
        self._timeout = configuration.getByName('ConnectTimeout')
        self._logger = Pool(ctx).getLogger('SpoolerLogger')
        self._logger.setDebugMode(True)

# Procedures called by SpoolerService
    def stop(self):
        if self.isStarted():
            self._disposed.set()
            self._thread.join()

    def start(self):
        if not self.isStarted():
            self._logger.logResource(INFO, 101)
            if self._isOffLine():
                self._logger.logResource(INFO, 102)
            else:
                self._logger.logResource(INFO, 103)
                self._disposed.clear()
                self._thread = Thread(target=self._run)
                self._thread.start()
                self._started()

    def isStarted(self):
        return self._thread is not None and self._thread.is_alive()

    def addListener(self, listener):
        self._listeners.append(listener)

    def removeListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

# Private methods
    def _run(self):
        try:
            connection = self._database.getConnection()
            jobs = self._database.getSpoolerJobs(connection)
            total = len(jobs)
            if total > 0:
                self._logger.logResource(INFO, 111, total)
                self._server = None
                count = self._send(connection, jobs)
                format = (count, total)
                self._logger.logResource(INFO, 112, format)
            else:
                self._logger.logResource(INFO, 113)
            connection.close()
        except UnoException as e:
            msg = self._logger.getMessage(114, e.Message)
            print(msg)
            self._logger.logMessage(INFO, msg)
        self._logger.logResource(INFO, 115)
        self._stopped()

    def _started(self):
        for listener in self._listeners:
            listener.started(self._event)

    def _stopped(self):
        for listener in self._listeners:
            listener.stopped(self._event)

    def _send(self, connection, jobs):
        print("MailSpooler._send()1 begin ****************************************")
        sf = getSimpleFile(self._ctx)
        uf = getUriFactory(self._ctx)
        sender = url = batch = None
        urls = ()
        count = 0
        for job in jobs:
            if self._disposed.is_set():
                print("MailSpooler._send() 2 break")
                break
            self._logger.logResource(INFO, 121, job)
            recipient = self._database.getRecipient(connection, job)
            print("MailSpooler._send() 3 batch %s" % recipient.BatchId)
            if batch != recipient.BatchId:
                try:
                    sender, urls, url = self._getBatchParameters(uf, sf, connection, recipient.BatchId, job, recipient.Index)
                    print("MailSpooler._send() 4 %s" % url.Main)
                except UnoException as e:
                    print("MailSpooler._send() 5 break %s" % e.Message)
                    self._database.setBatchState(connection, 2, recipient.BatchId)
                    msg = self._logger.getMessage(122, e.Message)
                    self._logger.logMessage(INFO, msg)
                    print("MailSpooler._send() 6 break %s" % msg)
                    continue
                else:
                    print("MailSpooler._send() 6")
                    batch = recipient.BatchId
            elif sender.Merge:
                descriptor = self._getDataDescriptor(sender, recipient.Index)
                url.merge(descriptor)
                print("MailSpooler._send() 7")
            if self._disposed.is_set():
                print("MailSpooler._send() 8 break")
                break
            state = self._getSendState(uf, sf, job, sender, recipient, urls, url)
            timestamp = getDateTime()
            self._database.setJobState(connection, state, timestamp, job)
            resource = 122 + state
            self._logger.logResource(INFO, resource, job)
            count += 1 if state == 1 else 0
            print("MailSpooler._send() 9")
        if self._server is not None:
            print("MailSpooler._send() 10 disconnect")
            self._server.disconnect()
            self._server = None
        print("MailSpooler._send() 11 end.........................")
        return count

    def _getSendState(self, uf, sf, job, sender, recipient, urls, url):
        state = 2
        body = MailTransferable(self._ctx, url.Main, True)
        mail = getMail(self._ctx, sender.Sender, recipient.Recipient, sender.Subject, body)
        if urls:
            state = self._sendMailWithAttachments(uf, sf, sender, mail, urls, job, recipient.Index)
        else:
            state = self._sendMail(mail)
        return state

    def _sendMailWithAttachments(self, uf, sf, sender, mail, urls, job, index):
        for url in urls:
            print("MailSpooler._sendMailWithAttachments() 1 %s" % url.Main)
            if sender.Merge and url.Merge:
                descriptor = self._getDataDescriptor(sender, index)
                url.merge(descriptor)
            mail.addAttachment(self._getAttachment(url))
        self._server.sendMailMessage(mail)
        return 1

    def _getAttachment(self, url):
        attachment = uno.createUnoStruct('com.sun.star.mail.MailAttachment')
        attachment.Data = MailTransferable(self._ctx, url.Main)
        attachment.ReadableName = url.Name
        return attachment

    def _sendMail(self, mail):
        self._server.sendMailMessage(mail)
        return 1

    def _getBatchParameters(self, uf, sf, connection, batch, job, index):
        urls = ()
        sender = self._database.getSender(connection, batch)
        if sender is None:
            raise self._getUnoException(131, job)
        self._checkUrl(sf, sender.Document, job, 132)
        uri = uf.parse(sender.Document)
        if sender.Merge:
            self._descriptor = self._getDescriptor(sender)
            descriptor = self._getDataDescriptor(sender, index)
        else:
            descriptor = False
        url = MailUrl(self._ctx, uri, sender.Merge, False, descriptor)
        urls = self._getMailUrls(uf, sf, connection, batch, job)
        server = self._database.getServer(connection, sender.Sender, self._timeout)
        self._checkServer(job, sender, server)
        return sender, urls, url

    def _getDataDescriptor(self, sender, index):
        bookmark = self._getBookmark(sender, index)
        self._descriptor['Selection'] = (bookmark, )
        descriptor = getPropertyValueSet(self._descriptor)
        return descriptor

    def _getBookmark(self, sender, index):
        connection = self._descriptor['ActiveConnection']
        format = (sender.Bookmark, sender.Table, sender.Identifier)
        bookmark = self._database.getBookmark(connection, format, index)
        return bookmark

    def _getMailUrls(self, uf, sf, connection, batch, job):
        urls = []
        attachments = self._database.getAttachments(connection, batch)
        for attachment in attachments:
            url = uf.parse(attachment)
            merge, pdf = self._getUrlMark(url)
            self._checkUrl(sf, url.UriReference, job, 141)
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

    def _checkUrl(self, sf, url, job, resource):
        if not sf.exists(url):
            format = (job, url)
            raise self._getUnoException(resource, format)

    def _checkServer(self, job, sender, server):
        if server is None:
            format = (job, sender.Sender)
            raise self._getUnoException(151, format)
        context = CurrentContext(server)
        authenticator = Authenticator(server)
        service = 'com.sun.star.mail.MailServiceProvider2'
        try:
            self._server = createService(self._ctx, service).create(SMTP)
            self._server.connect(context, authenticator)
        except UnoException as e:
            smtpserver = '%s:%s' % (server['ServerName'], server['Port'])
            format = (job, sender.Sender, smtpserver, e.Message)
            raise self._getUnoException(152, format)

    def _getUnoException(self, resource, format=None):
        e = UnoException()
        e.Message = self._logger.getMessage(resource, format)
        e.Context = self
        return e

    def _isOffLine(self):
        mode = getConnectionMode(self._ctx, *g_dns)
        return mode == OFFLINE

    def _getDescriptor(self, sender):
        datasource = self._getDocumentDataSource(sender.DataSource)
        connection = self._getDataSourceConnection(datasource)
        rowset = self._getRowSet(connection, sender.Table)
        rowset.execute()
        descriptor = {'ActiveConnection': connection,
                      'DataSourceName': sender.DataSource,
                      'Command': sender.Table,
                      'CommandType': TABLE,
                      'BookmarkSelection': False,
                      'Cursor': rowset.createResultSet()}
        return descriptor

        self._descriptor = self._getDescriptor(connection, sender, rowset)

    def _getDocumentDataSource(self, name):
        service = 'com.sun.star.sdb.DatabaseContext'
        dbcontext = createService(self._ctx, service)
        datasource = dbcontext.getByName(name)
        return datasource

    def _getDataSourceConnection(self, datasource):
        if datasource.IsPasswordRequired:
            handler = getInteractionHandler(self._ctx)
            connection = datasource.getIsolatedConnectionWithCompletion(handler)
        else:
            connection = datasource.getIsolatedConnection('', '')
        return connection

    def _getRowSet(self, connection, table):
        service = 'com.sun.star.sdb.RowSet'
        rowset = createService(self._ctx, service)
        rowset.ActiveConnection = connection
        rowset.CommandType = TABLE
        rowset.FetchSize = g_fetchsize
        rowset.Command = table
        return rowset


class MailUrl(unohelper.Base):
    def __init__(self, ctx, url, merge, pdf, descriptor=None):
        self._ctx = ctx
        self._merge = merge
        self._html = descriptor is not None
        self._pdf = pdf
        self._url = url.UriReference
        name = url.getPathSegment(url.PathSegmentCount -1)
        self._name = name
        self._title = name
        if self._isTemp():
            self._temp = getTempFile(ctx).Uri
            self._document = getDocument(ctx, self._url)
            if not self._merge:
                self._title = self._saveTempDocument()
            elif self._html:
                self.merge(descriptor)

    @property
    def Merge(self):
        return self._merge
    @property
    def Name(self):
        return self._title
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
        format = None
        if self._html:
            format = 'html'
        elif self._pdf:
            format = 'pdf'
        return saveTempDocument(self._document, self._temp, self._name, format)
