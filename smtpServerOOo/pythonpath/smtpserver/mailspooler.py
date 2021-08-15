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

from multiprocessing import Process, Queue, cpu_count
from threading import Thread
from threading import Event
import traceback
import time


class MailSpooler(unohelper.Base):
    def __init__(self, ctx, database):
        self._ctx = ctx
        self._database = database
        self._count = 0
        self._disposed = Event()
        self._batch = None
        self._server = None
        self._sender = None
        self._attachments = ()
        self._server = None
        self._process = None
        configuration = getConfiguration(ctx, g_identifier, False)
        self._timeout = configuration.getByName('ConnectTimeout')
        self._logger = Pool(ctx).getLogger('MailSpooler')
        self._logger.setDebugMode(True)

# Procedures called by MailServiceSpooler
    def stop(self):
        if self.isStarted():
            self._disposed.set()
            self._process.join()

    def start(self):
        try:
            print("MailSpooler.start(): 1 %s" % cpu_count())
            if not self.isStarted():
                logger = _getLogger(self._ctx, 'MailSpooler')
                _logMessage(logger, INFO, 101)
                if _isOffLine(self._ctx):
                    _logMessage(logger, INFO, 102)
                else:
                    print("MailSpooler.start(): 2")
                    connection = self._database.getConnection()
                    print("MailSpooler.start(): 3")
                    jobs = self._database.getSpoolerJobs(connection)
                    total = len(jobs)
                    print("MailSpooler.start(): 4")
                    if total > 0:
                        _logMessage(logger, INFO, 103, total)
                        print("MailSpooler.start(): 5")
                        self._disposed.clear()
                        configuration = getConfiguration(self._ctx, g_identifier, False)
                        timeout = configuration.getByName('ConnectTimeout')
                        arguments = (self._ctx, self._database, connection, logger, jobs, self._disposed, timeout)
                        self._process = Process(target=_mainProcess, args=arguments)
                        self._process.daemon = False
                        self._process.start()
                    else:
                        _logMessage(logger, INFO, 104)
                        print("MailSpooler.start(): 6")
                        connection.close()
                        _logMessage(logger, INFO, 105)
            print("MailSpooler.start(): 7")
        except Exception as e:
            msg = "MailSpooler.start(): Error: %s" % traceback.print_exc()
            print(msg)

    def isStarted(self):
        return self._process is not None and self._process.is_alive()


# Tasks Process:
def _mainProcess(ctx, database, connection, logger, jobs, disposed, timeout):
    try:
        print("MailSpooler._mainProcess()1 start ****************************************")
        sender = None
        attachments = ()
        url = ''
        count = 0
        server = None
        document = None
        batch = None
        for job in jobs:
            time.sleep(2)
            if disposed.is_set():
                print("MailSpooler._mainProcess() 2 canceled *******************************************")
                break
            _logMessage(logger, INFO, 121, job)
            print("MailSpooler._mainProcess() 3 %s" % job)
            time.sleep(2)
        if server is not None:
            print("MailSpooler._mainProcess() 4 disconnect")
            server.disconnect()
        if not connection.isClosed():
            print("MailSpooler._mainProcess() 5 close")
            #connection.close()
            print("MailSpooler._mainProcess() 6 close")
        _logMessage(logger, INFO, 105)
        print("MailSpooler._mainProcess() 7 stop *******************************************")
    except Exception as e:
        msg = "MailSpooler _mainProcess(): Error: %s" % traceback.print_exc()
        print(msg)


def _isOffLine(ctx):
    mode = getConnectionMode(ctx, *g_dns)
    return mode == OFFLINE

def _getLogger(ctx, name):
    logger = Pool(ctx).getLogger(name)
    logger.setDebugMode(True)
    return logger

def _logMessage(logger, level, resource, format=None):
    msg = logger.getMessage(resource, format)
    logger.logMessage(level, msg)








































# Private methods
def _run(self, jobs, database):
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
            #self._logMessage(INFO, 121, job)
            print("MailSpooler._run() 3")
            recipient = database.getRecipient(job)
            print("MailSpooler._run() 4")
            batch = recipient.BatchId
            print("MailSpooler._run() 5 %s" % batch)
            if self._batch != batch:
                try:
                    sender, document, attachments, url = self._canSetBatch(database, job, batch)
                except UnoException as e:
                    print("MailSpooler._run() 5 break")
                    database.setBatchState(batch, 2)
                    self._logger.logMessage(INFO, e.Message)
                    print("MailSpooler._run() 6 break")
                    break
            state = self._getSendState(ctx, job, sender, recipient, document, attachments, url)
            database.setJobState(job, state)
            resource = 110 + state
            self._logMessage(INFO, resource, job)
        if self._server is not None:
            print("MailSpooler._run() 7 disconnect")
            self._server.disconnect()
        #self.DataBase.dispose()
        print("MailSpooler._run() 8 canceled *******************************************")
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

def _canSetBatch(self, database, job, batch):
    url = ''
    print("MailSpooler._canSetBatch() 1")
    sender = database.getSender(batch)
    doc = sender.Document
    print("MailSpooler._canSetBatch() 2")
    self._checkUrl(doc, job, 131)
    print("MailSpooler._canSetBatch() 3")
    attachments = database.getAttachments(batch)
    print("MailSpooler._canSetBatch() 4")
    self._checkAttachments(job, attachments, 132)
    server = database.getServer(sender.Sender, self._timeout)
    self._checkServer(job, sender, server)
    print("MailSpooler._canSetBatch() 5")
    self._document = getDocument(self._ctx, doc)
    print("MailSpooler._canSetBatch() 6")
    if sender.DataSource is None:
        url = saveDocumentAs(self._ctx, self._document, 'html')
    self._batch = batch
    print("MailSpooler._canSetBatch() 7")
    return sender, attachments, url

def _getUrlContent(self, url):
    return getFileSequence(self._ctx, url)

def _checkUrl(self, url, job, resource):
    if not getSimpleFile(self._ctx).exists(url):
        print("MailSpooler._checkUrl() 1: %s" % url)
        format = (job, url)
        msg = self._logger.getMessage(resource, format)
        print("MailSpooler._checkUrl() 2: %s" % msg)
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

def _logMessage1(self, level, resource, format=None):
    msg = self._logger.getMessage(resource, format)
    self._logger.logMessage(level, msg)

def _dbworker1(ctx, database, logger, jobs, disposed, timeout):
    try:
        print("MailSpooler._run()1 begin ****************************************")
        sender = None
        attachments = ()
        url = ''
        count = 0
        server = None
        document = None
        batch = None
        for job in jobs:
            if disposed.is_set():
                print("MailSpooler._run() 2 break")
                break
            _logMessage(logger, INFO, 121, job)
            print("MailSpooler._run() 3 %s" % job)
            recipient = database.getRecipient(job)
            print("MailSpooler._run() 4 %s" % recipient)
            if batch != recipient.BatchId:
                try:
                    server, sender, document, attachments, url = _canSetBatch(ctx, logger, database, job, recipient.BatchId, timeout)
                    batch = recipient.BatchId
                    print("MailSpooler._run() 5")
                except UnoException as e:
                    print("MailSpooler._run() 6 break")
                    database.setBatchState(batch, 2)
                    _logMessage(logger, INFO, e.Message)
                    print("MailSpooler._run() 7 break")
                    break
            state = _getSendState(ctx, server, job, sender, recipient, document, attachments, url)
            database.setJobState(job, state)
            resource = 110 + state
            _logMessage(logger, INFO, resource, job)
        if server is not None:
            print("MailSpooler._run() 8 disconnect")
            server.disconnect()
        #self.DataBase.dispose()
        print("MailSpooler._run() 9 canceled *******************************************")
    except Exception as e:
        msg = "MailSpooler _run(): Error: %s" % traceback.print_exc()
        print(msg)

def _canSetBatch(ctx, logger, database, job, batch, timeout):
    url = ''
    server = None
    print("MailSpooler._canSetBatch() 1")
    sender = database.getSender(batch)
    doc = sender.Document
    print("MailSpooler._canSetBatch() 2")
    _checkUrl(ctx, logger, doc, job, 131)
    print("MailSpooler._canSetBatch() 3")
    attachments = database.getAttachments(batch)
    print("MailSpooler._canSetBatch() 4")
    _checkAttachments(ctx, logger, job, attachments, 132)
    config = database.getServer(sender.Sender, timeout)
    if _checkConfig(logger, sender, job, config):
        server = _getServer(ctx, logger, sender, job, config)
    print("MailSpooler._canSetBatch() 5")
    document = getDocument(ctx, doc)
    print("MailSpooler._canSetBatch() 6")
    if sender.DataSource is None:
        url = saveDocumentAs(ctx, document, 'html')
    print("MailSpooler._canSetBatch() 7")
    return server, sender, document, attachments, url

def _checkUrl(ctx, logger, url, job, resource):
    if not getSimpleFile(ctx).exists(url):
        print("MailSpooler._checkUrl() 1: %s" % url)
        format = (job, url)
        msg = logger.getMessage(resource, format)
        print("MailSpooler._checkUrl() 2: %s" % msg)
        raise UnoException(msg)

def _checkAttachments(ctx, logger, job, attachments, resource):
    for url in attachments:
        _checkUrl(ctx, logger, url, job, resource)

def _checkConfig(logger, sender, job, config):
    if config is None:
        format = (job, sender.Sender)
        msg = logger.getMessage(141, format)
        raise UnoException(msg)
    return True

def _getServer(ctx, logger, sender, job, config):
    context = CurrentContext(config)
    authenticator = Authenticator(config)
    service = 'com.sun.star.mail.MailServiceProvider2'
    try:
        server = createService(ctx, service).create(SMTP)
        server.connect(context, authenticator)
    except UnoException as e:
        smtpserver = '%(ServerName)s:%(Ports)' % config
        format = (job, sender.Sender, smtpserver, e.Message)
        msg = self._logger.getMessage(142, format)
        raise UnoException(msg)
    print("MailSpooler_isServer() ******************************")
    return server

def _getSendState(ctx, server, job, sender, recipient, document, attachments, url):
    print("MailSpooler._send() 1")
    if sender.DataSource is not None:
        # TODO: need to merge the document
        url = saveDocumentAs(ctx, document, 'html')
    print("MailSpooler._send() 2 %s - %s - %s - %s - %s - %s - %s - %s" % (sender.Sender, sender.Subject, sender.Document , sender.DataSource , sender.Query , recipient.Recipient , recipient.Identifier , attachments))
    mail = getMail(ctx, sender.Sender, recipient.Recipient, sender.Subject, url, True)
    if attachments:
        server.sendMailMessage(mail)
    else:
        server.sendMailMessage(mail)
    return 1
