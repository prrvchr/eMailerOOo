#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.util import XCloseListener

from com.sun.star.uno import Exception as UnoException

from com.sun.star.mail.MailServiceType import SMTP

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import getConfiguration
from unolib import createService
from unolib import getUrl

from .dbtools import getDataSource

from .database import DataBase
from .dataparser import DataParser

from .configuration import g_identifier

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
        self._error = None
        self._progress = 0.2
        self._configuration = getConfiguration(self.ctx, g_identifier, False)
        if not self._isInitialized():
            print("DataSource.__init__() 2")
            DataSource._Initialized = True
            initdatabase = self._addThread(Thread(target=self._initDataBase))
            initdatabase.start()
        print("DataSource.__init__() 3")

    _DataBase = None
    _Initialized = False
    _ThreadPool = []

    @property
    def DataBase(self):
        return DataSource._DataBase
    @DataBase.setter
    def DataBase(self, database):
        DataSource._DataBase = database

    def _isInitialized(self):
        return DataSource._Initialized

    def _addThread(self, thread):
        DataSource._ThreadPool.append(thread)
        return thread

    def _waitForThread(self):
        while len(DataSource._ThreadPool) > 1:
            thread = DataSource._ThreadPool.pop(0)
            thread.join()

    # XRestReplicator
    def cancel(self):
        self.canceled = True
        self.sync.set()
        self.join()

    def getSmtpConfig(self, email, progress, callback):
        arguments = (email, progress, callback)
        smtpconfig = self._addThread(Thread(target=self._getSmtpConfig, args=arguments))
        smtpconfig.start()

    def _getSmtpConfig(self, email, progress, callback):
        url = getUrl(self.ctx, self._configuration.getByName('IspDBUrl'))
        service = 'com.gmail.prrvchr.extensions.OAuth2OOo.OAuth2Service'
        request = createService(self.ctx, service)
        progress(10)
        offline = request.isOffLine(url.Server)
        progress(20)
        self._waitForThread()
        progress(40)
        user, servers = self.DataBase.getSmtpConfig(email)
        if len(servers) > 0:
            progress(100, 1)
        elif offline:
            progress(100, 2)
        else:
            progress(60)
            response = self._getIspdbConfig(request, url.Complete, user.getValue('Domain'))
            if response.IsPresent:
                progress(80)
                servers = self.DataBase.setSmtpConfig(response.Value)
                progress(100, 3)
            else:
                progress(100, 4)
        callback(user, servers, offline)

    def smtpConnect(self, context, authenticator, dialog):
        arguments = (context, authenticator, dialog)
        connect = self._addThread(Thread(target=self._smtpConnect, args=arguments))
        connect.start()

    def _smtpConnect(self, context, authenticator, dialog):
        dialog.updateProgress(25)
        print("DataSource._smtpConnect() 1")
        service = 'com.sun.star.mail.MailServiceProvider'
        server = createService(self.ctx, service).create(SMTP)
        print("DataSource._smtpConnect() 2")
        dialog.updateProgress(50)
        try:
            server.connect(context, authenticator)
        except UnoException as e:
            print("DataSource._smtpConnect() ERROR: %s" % e.Message)
            dialog.updateProgress(100, 2)
        else:
            dialog.updateProgress(75)
            if server.isConnected():
                dialog.updateProgress(100, 1)
                format = (server.isConnected(), server.getSupportedAuthenticationTypes())
                print("DataSource._smtpConnect() 3 isConnected: %s - %s" % format)
                server.disconnect()
                dialog.callBack()
            else:
                dialog.updateProgress(100, 2)
        print("DataSource._smtpConnect() 4")

    def _getIspdbConfig(self, request, url, domain):
        parameter = uno.createUnoStruct('com.sun.star.auth.RestRequestParameter')
        parameter.Method = 'GET'
        parameter.Url = '%s%s' % (url, domain)
        parameter.NoAuth = True
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
