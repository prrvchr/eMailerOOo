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

from com.sun.star.uno import Exception as UnoException


from .unotool import createService
from .unotool import getConfiguration
from .unotool import getConnectionMode
from .unotool import getFileSequence
from .unotool import getSimpleFile

from smtpserver import getMail
from smtpserver import getDocument
from smtpserver import saveDocumentAs
from smtpserver import Authenticator
from smtpserver import CurrentContext

from .logger import Logger
from .logger import Pool

from .configuration import g_dns
from .configuration import g_identifier

from threading import Thread
from threading import Event
import traceback


class MailSpooler(unohelper.Base):
    def __init__(self, ctx, database):
        self._ctx = ctx
        self.DataBase = database
        self._count = 0
        self._disposed = Event()
        self._batch = None
        self._server = None
        self._sender = None
        self._attachments = ()
        self._server = None
        self._thread = None
        configuration = getConfiguration(ctx, g_identifier, False)
        self._timeout = configuration.getByName('ConnectTimeout')
        self._logger = Pool(ctx).getLogger('MailSpooler')
        self._logger.setDebugMode(True)

# Procedures called by MailServiceSpooler
    def stop(self):
        if self.isStarted():
            self._disposed.set()
            self._thread.join()

    def start(self):
        if not self.isStarted():
            if self._isOffLine():
                self._logMessage(INFO, 101)
            else:
                jobs = self.DataBase.getSpoolerJobs(0)
                if jobs:
                    self._disposed.clear()
                    self._thread = Thread(target=self._run, args=jobs)
                    self._thread.start()
                #else:
                #    self.DataBase.dispose()

    def isStarted(self):
        return self._thread is not None and self._thread.is_alive()

# Private methods
    def _run(self, *jobs):
        try:
            print("MailSpooler._run()1 begin ****************************************")
            sender = None
            attachments = ()
            url = ''
            count = 0
            self._server = None
            for job in jobs:
                if self._disposed.is_set():
                    print("MailSpooler._run() 2 break")
                    break
                self._logMessage(INFO, 121, job)
                recipient = self.DataBase.getRecipient(job)
                batch = recipient.BatchId
                if self._batch != batch:
                    try:
                        sender, attachments, url = self._canSetBatch(job, batch)
                    except UnoException as e:
                        self.DataBase.setBatchState(batch, 2)
                        self._logger.logMessage(INFO, e.Message)
                        print("MailSpooler._run() 3 break")
                        break
                state = self._getSendState(job, sender, recipient, attachments, url)
                self.DataBase.setJobState(job, state)
                resource = 110 + state
                self._logMessage(INFO, resource, job)
            if self._server is not None:
                print("MailSpooler._run() 4 disconnect")
                self._server.disconnect()
            #self.DataBase.dispose()
            print("MailSpooler._run() 5 canceled *******************************************")
        except Exception as e:
            msg = "MailSpooler _run(): Error: %s" % traceback.print_exc()
            print(msg)

    def _getSendState(self, job, sender, recipient, attachments, url):
        print("MailSpooler._send() 1")
        if sender.DataSource is not None:
            # TODO: need to merge the document
            url = saveDocumentAs(self._ctx, self._document, 'html')
        print("MailSpooler._send() 2 %s - %s - %s - %s - %s - %s - %s - %s" % (sender.Sender, sender.Subject, sender.Document , sender.DataSource , sender.Query , recipient.Recipient , recipient.Identifier , attachments))
        mail = getMail(self._ctx, sender.Sender, recipient.Recipient, sender.Subject, url, True)
        if attachments:
            self._sendMailWithAttachments(mail, attachments)
        else:
            self._sendMail(mail)



        print("MailSpooler._send() 4")
        return 1

    def _sendMailWithAttachments(self, recipient, body):
        print("MailSpooler._sendMailWithAttachments() 1")
        pass


    def _sendMail(self, mail):
        print("MailSpooler._sendMail() 1")
        self._server.sendMailMessage(mail)

    def _canSetBatch(self, job, batch):
        url = ''
        sender = self.DataBase.getSender(batch)
        doc = sender.Document
        self._checkUrl(doc, job, 131)
        attachments = self.DataBase.getAttachments(batch)
        self._checkAttachments(job, attachments, 132)
        server = self.DataBase.getServer(sender.Sender, self._timeout)
        self._checkServer(job, sender, server)
        self._document = getDocument(self._ctx, doc)
        if sender.DataSource is None:
            url = saveDocumentAs(self._ctx, self._document, 'html')
        self._batch = batch
        return sender, attachments, url

    def _getUrlContent(self, url):
        return getFileSequence(self._ctx, url)

    def _checkUrl(self, url, job, resource):
        if not getSimpleFile(self._ctx).exists(url):
            format = (job, url)
            msg = self._logger.getMessage(resource, format)
            raise UnoException(msg)

    def _checkAttachments(self, job, attachments, resource):
        for url in attachments:
            self._checkUrl(url, job, resource)

    def _checkServer(self, job, sender, server):
        if server is None:
            format = (job, sender.Sender)
            msg = self._logger.getMessage(141, format)
            raise UnoException(msg)
        context = CurrentContext(server)
        authenticator = Authenticator(server)
        service = 'com.sun.star.mail.MailServiceProvider2'
        try:
            self._server = createService(self._ctx, service).create(SMTP)
            self._server.connect(context, authenticator)
        except UnoException as e:
            smtpserver = '%s:%s' % (server['ServerName'], server['Port'])
            format = (job, sender.Sender, smtpserver, e.Message)
            msg = self._logger.getMessage(142, format)
            raise UnoException(msg)
        print("MailSpooler_isServer() ******************************")

    def _logMessage(self, level, resource, format=None):
        msg = self._logger.getMessage(resource, format)
        self._logger.logMessage(level, msg)

    def _isOffLine(self):
        mode = getConnectionMode(self._ctx, *g_dns)
        return mode == OFFLINE
