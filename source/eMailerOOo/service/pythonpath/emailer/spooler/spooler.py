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

from com.sun.star.uno import Exception as UnoException

from ..mailerlib import MailTransferable

from ..listener import TerminateListener

from ..unotool import getConnectionMode
from ..unotool import getResourceLocation
from ..unotool import getSimpleFile
from ..unotool import getTempFile

from ..mailertool import getDataSource
from ..mailertool import getMailMessage
from ..mailertool import getMailService
from ..mailertool import getMessageImage

from ..logger import getLogger
from ..logger import RollerHandler

from ..configuration import g_dns
from ..configuration import g_extension
from ..configuration import g_identifier
from ..configuration import g_spoolerlog
from ..configuration import g_logo
from ..configuration import g_logourl

from .mailer import Mailer

from .mailer import getUnoException

from threading import Thread
from threading import Lock
from threading import Event
import traceback


class Spooler():
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

    def viewJob(self, jobid):
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('viewJobs')
        if self._isInitialized():
            self.DataSource.waitForDataBase()
            return self._viewJob(jobid)

# Private methods
    def _isInitialized(self):
        return self.DataSource is not None

    def _initialize(self, method):
        self.DataSource = getDataSource(self._ctx, method, ' ', self._logError)

    def _logError(self, method, code, *args):
        self._logger.logprb(SEVERE, 'MailSpooler', method, code, *args)

    def _started(self):
        event = uno.createUnoStruct('com.sun.star.lang.EventObject')
        for listener in self._listeners:
            listener.started(event)

    def _stopped(self):
        event = uno.createUnoStruct('com.sun.star.lang.EventObject')
        for listener in self._listeners:
            listener.stopped(event)

    def _run(self):
        handler = RollerHandler(self._ctx, self._logger.Name)
        self._logger.addRollerHandler(handler)
        self._logger.logprb(INFO, 'MailSpooler', 'start()', 1001)
        if not self._isInitialized():
            # FIXME: We need to check the configuration
            self._initialize('start')
        if self._isInitialized():
            # FIXME: Configuration has been checked we can continue
            if self._isOffLine():
                self._logger.logprb(INFO, 'MailSpooler', 'start()', 1002)
            else:
                self._logger.logprb(INFO, 'MailSpooler', 'start()', 1003)
                try:
                    self.DataSource.waitForDataBase()
                    connection = self.DataSource.DataBase.getConnection()
                    jobs, total = self.DataSource.DataBase.getSpoolerJobs(connection)
                    if total > 0:
                        self._logger.logprb(INFO, 'MailSpooler', 'start()', 1011, total)
                        count = self._sendMails(connection, jobs)
                        self._logger.logprb(INFO, 'MailSpooler', 'stop()', 1012, count, total)
                    else:
                        self._logger.logprb(INFO, 'MailSpooler', 'start()', 1013)
                    connection.close()
                except UnoException as e:
                    self._logger.logprb(SEVERE, 'MailSpooler', 'start()', 1014, e.Message)
                except Exception as e:
                    msg = "Error: %s" % traceback.format_exc()
                    self._logger.logprb(SEVERE, 'MailSpooler', 'start()', 1014, msg)
        self._logger.logprb(INFO, 'MailSpooler', 'stop()', 1015)
        self._logger.removeRollerHandler(handler)
        self._stopped()

    def _viewJob(self, job):
        handler = RollerHandler(self._ctx, self._logger.Name)
        self._logger.addRollerHandler(handler)
        self._logger.logprb(INFO, 'MailSpooler', '_viewJobs()', 1051, job)
        connection = self.DataSource.DataBase.getConnection()
        mailer = Mailer(self._ctx, self.DataSource.DataBase, self._logger)
        mail, new = self._getMail(connection, job, mailer)
        url = '%s/Email.eml' % getTempFile(self._ctx).Uri
        output = getSimpleFile(self._ctx).openFileWrite(url)
        output.writeBytes(uno.ByteSequence(mail.asBytes()))
        output.flush()
        output.closeOutput()
        mailer.dispose()
        connection.close()
        self._logger.logprb(INFO, 'MailSpooler', '_viewJobs()', 1052, job)
        self._logger.removeRollerHandler(handler)
        return url

    def _sendMails(self, connection, jobs):
        mailer = Mailer(self._ctx, self.DataSource.DataBase, self._logger, True)
        server = None
        count = 0
        for job in jobs:
            self._logger.logprb(INFO, 'MailSpooler', '_sendMails()', 1021, job)
            if self._stop.is_set():
                self._logger.logprb(INFO, 'MailSpooler', '_sendMails()', 1022, job)
                break
            mail, new = self._getMail(connection, job, mailer)
            if self._stop.is_set():
                self._logger.logprb(INFO, 'MailSpooler', '_sendMails()', 1022, job)
                break
            if new:
                server = self._getServer(SMTP, mailer.getUser(), server)
            server.sendMailMessage(mail)
            self.DataSource.DataBase.updateRecipient(connection, 1, mail.MessageId, job)
            self._logger.logprb(INFO, 'MailSpooler', '_sendMails()', 1023, job)
            count += 1
        self._dispose(server, mailer)
        return count

    def _getMail(self, connection, job, mailer):
        recipient = self.DataSource.DataBase.getRecipient(connection, job)
        new = mailer.isNewBatch(recipient.BatchId)
        if new:
            self._initMailer(connection, job, mailer, recipient)
        elif mailer.Merge:
            mailer.merge(recipient.Filter)
        body = MailTransferable(self._ctx, mailer.getBodyUrl(), True, True)
        mail = getMailMessage(self._ctx, mailer.Sender, recipient.Recipient, mailer.getSubject(), body)
        mailer.addAttachments(mail, recipient.Filter)
        return mail, new

    def _initMailer(self, connection, job, mailer, recipient):
        metadata = self.DataSource.DataBase.getMailer(connection, recipient.BatchId)
        attachments = self.DataSource.DataBase.getAttachments(connection, recipient.BatchId)
        mailer.setBatch(recipient.BatchId, metadata, attachments, job, recipient.Filter)
        if mailer.needThreadId():
            self._createThreadId(connection, mailer, recipient.BatchId)

    def _createThreadId(self, connection, mailer, batch):
        server = self._getServer(IMAP, mailer.getUser())
        folder = server.getSentFolder()
        if server.hasFolder(folder):
            subject = mailer.getSubject(False)
            message = self._getThreadMessage(mailer, batch, subject)
            body = MailTransferable(self._ctx, message, True)
            mail = getMailMessage(self._ctx, mailer.Sender, mailer.Sender, subject, body)
            server.uploadMessage(folder, mail)
            mailer.setThreadId(connection, batch, mail.MessageId)
        server.disconnect()

    def _getThreadMessage(self, mailer, batch, subject):
        title = self._logger.resolveString(1031, g_extension, batch, mailer.Query)
        message = self._logger.resolveString(1032)
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
''' % (g_extension, logo, g_logourl, title, message, subject,
        document, mailer.Document, mailer.getDocumentTitle(),
        files, attachments)

    def _dispose(self, server, mailer):
        if server is not None:
            server.disconnect()
        mailer.dispose()

    def _getServer(self, mailtype, user, server=None):
        if server is not None:
            server.disconnect()
        server = self._getMailServer(mailtype, user)
        return server

    def _checkUrl(self, sf, url, job, resource):
        if not sf.exists(url):
            raise getUnoException(self._logger, self, resource, job, url)

    def _getMailServer(self, mailtype, user):
        context = user.getConnectionContext(mailtype)
        authenticator = user.getAuthenticator(mailtype)
        domain = user.getServerDomain(mailtype)
        try:
            server = getMailService(self._ctx, mailtype.value, domain)
            server.connect(context, authenticator)
        except UnoException as e:
            host = '%s:%s' % (context.getValueByName('ServerName'), context.getValueByName('Port'))
            raise getUnoException(self._logger, self, 1041, job, sender.Sender, host, e.Message)
        return server

    def _isOffLine(self):
        mode = getConnectionMode(self._ctx, *g_dns)
        return mode == OFFLINE

