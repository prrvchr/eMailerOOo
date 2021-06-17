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

from com.sun.star.ucb.ConnectionMode import OFFLINE

from .unotool import getConnectionMode
from .unotool import getFileSequence
from .unotool import getSimpleFile

from smtpserver import getDocument
from smtpserver import saveDocumentAs

from .logger import Logger

from .configuration import g_dns

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
        self._sender = None
        self._attachments = ()
        self._server = None
        self._thread = None
        self._logger = Logger(ctx, 'MailSpooler', 'MailSpooler')
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
            for job in jobs:
                if self._disposed.is_set():
                    print("MailSpooler._run() 2 break")
                    break
                self._count = 0
                state = self._getSendState(job)
                self.DataBase.setJobState(job, state)
                resource = 110 + state
                self._logMessage(INFO, resource, job)
            #self.DataBase.dispose()
            print("MailSpooler._run() 3 canceled *******************************************")
        except Exception as e:
            msg = "MailSpooler _run(): Error: %s" % traceback.print_exc()
            print(msg)

    def _getSendState(self, job):
        print("MailSpooler._send() 1")
        self._logMessage(INFO, 121, job)
        recipient = self.DataBase.getRecipient(job)
        batch = recipient.BatchId
        if self._batch != batch:
            if not self._canSetBatch(job, batch):
                return 2
        if self._sender.DataSource is not None:
            # TODO: need to merge the document
            self._url = saveDocumentAs(self._ctx, self._document, 'html')

        print("MailSpooler._send() 2 %s - %s - %s - %s - %s - %s - %s - %s" % (self._sender.Sender, self._sender.Subject, self._sender.Document , self._sender.DataSource , self._sender.Query , recipient.Recipient , recipient.Identifier , self._attachments))
        print("MailSpooler._send() 3 %s - %s - %s - %s - %s - %s - %s" % (self._server.Server, self._server.Port, self._server.Connection , self._server.Authentication , self._server.LoginMode , self._server.LoginName , self._server.Password))

        print("MailSpooler._send() 4")
        return 1


    def _canSetBatch(self, job, batch):
        self._sender = self.DataBase.getSender(batch)
        url = self._sender.Document
        if not self._isUrl(url):
            format = (job, url)
            self._logMessage(INFO, 131, format)
            return False
        self._attachments = self.DataBase.getAttachments(batch)
        if not self._isAttachments(job):
            return False
        self._server = self.DataBase.getServer(self._sender.Sender)
        if not self._isServer(job, self._sender.Sender):
            return False
        self._document = getDocument(self._ctx, self._sender.Document)
        if self._sender.DataSource is None:
            self._url = saveDocumentAs(self._ctx, self._document, 'html')
        self._batch = batch
        return True

    def _getUrlContent(self, url):
        return getFileSequence(self._ctx, url)

    def _isUrl(self, url):
        return getSimpleFile(self._ctx).exists(url)

    def _isAttachments(self, job):
        for url in self._attachments:
            if not self._isUrl(url):
                format = (job, url)
                self._logMessage(INFO, 141, format)
                return False
        return True

    def _isServer(self, job, sender):
        if self._server is None:
            format = (job, sender)
            self._logMessage(INFO, 151, format)
            return False
        return True

    def _logMessage(self, level, resource, format=None):
        msg = self._logger.getMessage(resource, format)
        self._logger.logMessage(level, msg)

    def _isOffLine(self):
        mode = getConnectionMode(self._ctx, *g_dns)
        return mode == OFFLINE
