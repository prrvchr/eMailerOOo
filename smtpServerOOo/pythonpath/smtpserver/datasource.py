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

from com.sun.star.util import XCloseListener
from com.sun.star.datatransfer import XTransferable

from com.sun.star.uno import Exception as UnoException

from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.ucb.ConnectionMode import OFFLINE
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService
from unolib import getConnectionMode
from unolib import getUrl

from .dbtools import getDataSource

from .database import DataBase
from .dataparser import DataParser

from .configuration import g_identifier

from .logger import setDebugMode
from .logger import logMessage
from .logger import getMessage

from multiprocessing import Process
from threading import Thread
#from threading import Condition

import traceback
import time


class DataSource(unohelper.Base,
                 XCloseListener):
    def __init__(self, ctx):
        print("DataSource.__init__() 1")
        self.ctx = ctx
        self._dbname = 'SmtpServer'
        self._config = None
        if not self._isInitialized():
            print("DataSource.__init__() 2")
            DataSource._Init = Thread(target=self._initDataBase)
            DataSource._Init.start()
        print("DataSource.__init__() 3")

    _Init = None
    _DataBase = None

    @property
    def DataBase(self):
        return DataSource._DataBase
    @DataBase.setter
    def DataBase(self, database):
        DataSource._DataBase = database

    def _isInitialized(self):
        return DataSource._Init is not None

    # XRestReplicator
    def cancel(self):
        self.canceled = True
        self.sync.set()
        self.join()

    def saveUser(self, *args):
        self.DataBase.mergeUser(*args)

    def saveServer(self, new, provider, host, port, server):
        if new:
            self.DataBase.mergeProvider(provider)
            self.DataBase.mergeServer(provider, server)
        else:
            self.DataBase.updateServer(host, port, server)

    def waitForDataBase(self):
        DataSource._Init.join()

    def getConfig(self, email):
        self.waitForDataBase()
        config = self.DataBase.getConfig(email)
        return config

    def getSmtpConfig(self, *args):
        config = Thread(target=self._getSmtpConfig, args=args)
        config.start()

    def _getSmtpConfig(self, email, url, progress, callback):
        progress(5)
        url = getUrl(self.ctx, url)
        progress(10)
        mode = getConnectionMode(self.ctx, url.Server)
        progress(20)
        self.waitForDataBase()
        progress(40)
        user, servers = self.DataBase.getSmtpConfig(email)
        if len(servers) > 0:
            progress(100, 1)
        elif mode == OFFLINE:
            progress(100, 2)
        else:
            progress(60)
            service = 'com.gmail.prrvchr.extensions.OAuth2OOo.OAuth2Service'
            request = createService(self.ctx, service)
            response = self._getIspdbConfig(request, url.Complete, user.getValue('Domain'))
            if response.IsPresent:
                progress(80)
                servers = self.DataBase.setSmtpConfig(response.Value)
                progress(100, 3)
            else:
                progress(100, 4)
        callback(user, servers, mode)

    def smtpConnect(self, *args):
        setDebugMode(self.ctx, True)
        connect = Thread(target=self._smtpConnect, args=args)
        connect.start()

    def _smtpConnect(self, context, authenticator, progress, callback):
        progress(0)
        step = 2
        progress(5)
        service = 'com.sun.star.mail.MailServiceProvider2'
        progress(25)
        server = createService(self.ctx, service).create(SMTP)
        progress(50)
        try:
            server.connect(context, authenticator)
        except UnoException as e:
            progress(100)
        else:
            progress(75)
            if server.isConnected():
                server.disconnect()
                step = 4
                progress(100)
            else:
                progress(100)
        setDebugMode(self.ctx, False)
        callback(step)

    def smtpSend(self, *args):
        setDebugMode(self.ctx, True)
        send = Thread(target=self._smtpSend, args=args)
        send.start()

    def _smtpSend(self, context, authenticator, sender, recipient, subject, message, progress, callback):
        step = 3
        progress(5)
        service = 'com.sun.star.mail.MailServiceProvider2'
        progress(25)
        server = createService(self.ctx, service).create(SMTP)
        progress(50)
        try:
            server.connect(context, authenticator)
        except UnoException as e:
            print("DataSoure._smtpSend() 1 Error: %s" % e.Message)
            progress(100)
        else:
            progress(75)
            if server.isConnected():
                service = 'com.sun.star.mail.MailMessage2'
                body = MailTransferable(self.ctx, message)
                arguments = (recipient, sender, subject, body)
                mail = createService(self.ctx, service, *arguments)
                print("DataSoure._smtpSend() 2: %s - %s" % (type(mail), mail))
                try:
                    server.sendMailMessage(mail)
                except UnoException as e:
                    print("DataSoure._smtpSend() 3 Error: %s - %s" % (e.Message, traceback.print_exc()))
                else:
                    step = 5
                server.disconnect()
                progress(100)
            else:
                progress(100)
        setDebugMode(self.ctx, False)
        callback(step)

    def _getIspdbConfig(self, request, url, domain):
        parameter = uno.createUnoStruct('com.sun.star.auth.RestRequestParameter')
        parameter.Method = 'GET'
        parameter.Url = '%s%s' % (url, domain)
        parameter.NoAuth = True
        #parameter.NoVerify = True
        response = request.getRequest(parameter, DataParser()).execute()
        return response

    def _initDataBase(self):
        self.DataBase = DataBase(self.ctx, self._dbname)
        self.DataBase.addCloseListener(self)

    # XCloseListener
    def queryClosing(self, source, ownership):
        self.DataBase.shutdownDataBase()
        msg = "DataBase  '%s' closing ... Done" % self._dbname
        logMessage(self.ctx, INFO, msg, 'DataSource', 'queryClosing()')
        print(msg)
    def notifyClosing(self, source):
        pass


class MailTransferable(unohelper.Base,
                       XTransferable):
    def __init__(self, ctx, body):
        print("MailTransferable.__init__() 1")
        self.ctx = ctx
        self._body = body
        self._html = False
        print("MailTransferable.__init__() 2")

    # XTransferable
    def getTransferData(self, flavor):
        if flavor.MimeType == "text/plain;charset=utf-16":
            print("MailTransferable.getTransferData() 1")
            data = self._body
        elif flavor.MimeType == "text/html;charset=utf-8":
            print("MailTransferable.getTransferData() 2")
            data = ''
        else:
            print("MailTransferable.getTransferData() 3")
            data = ''
        return data

    def getTransferDataFlavors(self):
        flavor = uno.createUnoStruct('com.sun.star.datatransfer.DataFlavor')
        if self._html:
            flavor.MimeType = 'text/html;charset=utf-8'
            flavor.HumanPresentableName = 'HTML-Documents'
        else:
            flavor.MimeType = 'text/plain;charset=utf-16'
            flavor.HumanPresentableName = 'Unicode text'
        print("MailTransferable.getTransferDataFlavors() 1")
        return (flavor,)

    def isDataFlavorSupported(self, flavor):
        support = flavor.MimeType == 'text/plain;charset=utf-16' or flavor.MimeType == 'text/html;charset=utf-8'
        print("MailTransferable.isDataFlavorSupported() 1 %s" % support)
        return support
