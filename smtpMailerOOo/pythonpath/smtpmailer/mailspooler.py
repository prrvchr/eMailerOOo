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
            jobs = self._database.getSpoolerJobs(connection)
            print("MailSpooler._run() 2")
            total = len(jobs)
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
        sf = getSimpleFile(self._ctx)
        uf = getUriFactory(self._ctx)
        server = config = datasource = sender = url = batchid = None
        urls = ()
        count = 0
        for job in jobs:
            if self._disposed.is_set():
                print("MailSpooler._send() 2 break")
                break
            self._logger.logResource(INFO, 121, job)
            recipient = self._database.getRecipient(connection, job)
            print("MailSpooler._send() 3 batch %s" % recipient.BatchId)
            if batchid != recipient.BatchId:
                try:
                    newsender = self._database.getSender(connection, recipient.BatchId)
                    sender, datasource = self._getSender(sender, newsender, datasource, sf, job)
                    print("MailSpooler._send() 4 %s" % (sender, ))
                    if self._disposed.is_set():
                        print("MailSpooler._send() 3 break")
                        break
                    self._disposeUrls(urls, url)
                    attachments = self._database.getAttachments(connection, recipient.BatchId)
                    urls, url = self._getUrls(uf, sf, attachments, sender, datasource, job, recipient.Index)
                    if self._disposed.is_set():
                        print("MailSpooler._send() 4 break")
                        break
                    newconfig = self._database.getServer(connection, sender.Sender, self._timeout)
                    server, config = self._getServer(server, config, newconfig, sender, job)
                    print("MailSpooler._send() 4 %s" % url.Main)
                except UnoException as e:
                    print("MailSpooler._send() 5 break %s" % e.Message)
                    self._database.setJobState(connection, 2, job)
                    msg = self._logger.getMessage(122, e.Message)
                    self._logger.logMessage(INFO, msg)
                    print("MailSpooler._send() 6 break %s" % msg)
                    continue
                else:
                    print("MailSpooler._send() 6")
                    batchid = recipient.BatchId
            elif sender.Merge:
                descriptor = self._getDataDescriptor(datasource, sender, recipient.Index)
                url.merge(descriptor)
                print("MailSpooler._send() 7")
            mail = self._getMail(sender, recipient, url)
            self._addAttachments(mail, urls, datasource, sender, recipient.Index)
            if self._disposed.is_set():
                print("MailSpooler._send() 8 break")
                break
            server.sendMailMessage(mail)
            self._database.setJobState(connection, 1, job)
            self._logger.logResource(INFO, 123, job)
            count += 1
            print("MailSpooler._send() 9")
        self._dispose(server, datasource, urls, url)
        print("MailSpooler._send() 11 end.........................")
        return count

    def _dispose(self, server, datasource, urls, url):
        self._disconnect(server)
        self._disposeDataSource(datasource)
        self._disposeUrls(urls, url)

    def _disconnect(self, server):
        if server is not None:
            server.disconnect()

    def _disposeDataSource(self, datasource):
        if datasource is not None:
            datasource['ActiveConnection'].close()

    def _disposeUrls(self, urls, url):
        if url is not None:
            url.dispose()
        for url in urls:
            url.dispose()

    def _getMail(self, sender, recipient, url):
        body = MailTransferable(self._ctx, url.Main, True)
        mail = getMail(self._ctx, sender.Sender, recipient.Recipient, sender.Subject, body)
        return mail

    def _addAttachments(self, mail, urls, datasource, sender, index):
        for url in urls:
            print("MailSpooler._addAttachments() %s - %s" % (url.Name, url.Main))
            if sender.Merge and url.Merge:
                descriptor = self._getDataDescriptor(datasource, sender, index)
                url.merge(descriptor)
            mail.addAttachment(self._getAttachment(url))

    def _getAttachment(self, url):
        attachment = uno.createUnoStruct('com.sun.star.mail.MailAttachment')
        attachment.Data = MailTransferable(self._ctx, url.Main)
        attachment.ReadableName = url.Name
        return attachment

    def _getSender(self, old, new, datasource, sf, job):
        self._checkSender(sf, new, job)
        if self._isNewDataSource(old, new, datasource):
            self._disposeDataSource(datasource)
            datasource = self._getDescriptor(new)
        return new, datasource

    def _checkSender(self, sf, sender, job):
        if sender is None:
            raise self._getUnoException(131, job)
        self._checkUrl(sf, sender.Document, job, 132)

    def _isNewDataSource(self, old, new, datasource):
        if not new.Merge:
            isnew = False
        elif old is None or datasource is None:
            isnew = True
        else:
            isnew =  any((old.DataSource != new.DataSource,
                          old.Table != new.Table,
                          old.Identifier != new.Identifier,
                          old.Bookmark != new.Bookmark))
        return isnew

    def _getUrls(self, uf, sf, attachments, sender, datasource, job, index):
        descriptor = False
        if sender.Merge:
            descriptor = self._getDataDescriptor(datasource, sender, index)
        uri = uf.parse(sender.Document)
        url = MailUrl(self._ctx, uri, sender.Merge, False, descriptor)
        urls = self._getMailUrl(uf, sf, attachments, job)
        return urls, url

    def _getServer(self, server, old, new, sender, job):
        if old != new:
            self._disconnect(server)
            server = self._getSmtpServer(sender, new, job)
        return server, new

    def _getDataDescriptor(self, datasource, sender, index):
        connection = datasource['ActiveConnection']
        bookmark = self._getBookmark(connection, sender, index)
        datasource['Selection'] = (bookmark, )
        descriptor = getPropertyValueSet(datasource)
        return descriptor

    def _getBookmark(self, connection, sender, index):
        format = (sender.Bookmark, sender.Table, sender.Identifier)
        bookmark = self._database.getBookmark(connection, format, index)
        return bookmark

    def _getMailUrl(self, uf, sf, attachments, job):
        urls = []
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

    def _getSmtpServer(self, sender, config, job):
        if config is None:
            format = (job, sender.Sender)
            raise self._getUnoException(151, format)
        context = CurrentContext(config)
        authenticator = Authenticator(config)
        service = 'com.sun.star.mail.MailServiceProvider2'
        try:
            server = createService(self._ctx, service).create(SMTP)
            server.connect(context, authenticator)
        except UnoException as e:
            smtpserver = '%s:%s' % (server['ServerName'], server['Port'])
            format = (job, sender.Sender, smtpserver, e.Message)
            raise self._getUnoException(152, format)
        return server

    def _getUnoException(self, resource, format=None):
        e = UnoException()
        e.Message = self._logger.getMessage(resource, format)
        e.Context = self
        return e

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
