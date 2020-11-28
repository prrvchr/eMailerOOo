#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.util import XCloseListener

from com.sun.star.uno import Exception as UnoException

from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.ucb.ConnectionMode import OFFLINE
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import getConfiguration
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

from collections import OrderedDict

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
        self._configuration = getConfiguration(self.ctx, g_identifier, False)
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

    def getSmtpConfig(self, *args):
        config = Thread(target=self._getSmtpConfig, args=args)
        config.start()

    def _getSmtpConfig(self, email, progress, callback):
        progress(5)
        url = getUrl(self.ctx, self._configuration.getByName('IspDBUrl'))
        progress(10)
        mode = getConnectionMode(self.ctx, url.Server)
        progress(20)
        DataSource._Init.join()
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
            print("DataSoure._smtpConnect() 1 Error: %s" % e.Message)
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

    def _smtpSend(self, context, authenticator, recipient, object, message, progress, callback):
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
                server.disconnect()
                step = 5
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
        self.DataBase = DataBase(self.ctx, 'SmtpServer')
        self.DataBase.addCloseListener(self)

    # XCloseListener
    def queryClosing(self, source, ownership):
        #if self.DataSource.is_alive():
        #    self.DataSource.cancel()
        #    self.DataSource.join()
        #self.deregisterInstance(self.Scheme, self.Plugin)
        #self.DataBase.shutdownDataBase(self.DataSource.fullPull)
        msg = "DataSource queryClosing: Scheme: %s ... Done" % 'SmtpServer'
        logMessage(self.ctx, INFO, msg, 'DataSource', 'queryClosing()')
        print(msg)
    def notifyClosing(self, source):
        pass
