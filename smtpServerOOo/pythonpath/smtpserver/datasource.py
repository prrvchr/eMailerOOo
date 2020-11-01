#!
# -*- coding: utf_8 -*-

#from __futur__ import absolute_import

import uno
import unohelper

from com.sun.star.util import XCloseListener

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
from threading import Thread
from threading import Condition
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
        if self._initializeDataBase():
            print("DataSource.__init__() 2")
            self.InitThread = Thread(target=self._initDataBase)
            self.InitThread.start()
        print("DataSource.__init__() 3")

    _DataBase = None
    _InitThread = None

    @property
    def InitThread(self):
        return DataSource._InitThread
    @InitThread.setter
    def InitThread(self, thread):
        DataSource._InitThread = thread
    @property
    def DataBase(self):
        return DataSource._DataBase
    @DataBase.setter
    def DataBase(self, database):
        DataSource._DataBase = database

    def _initializeDataBase(self):
        return self.InitThread is None

    # XRestReplicator
    def cancel(self):
        self.canceled = True
        self.sync.set()
        self.join()

    def getSmtpConfig(self, email, progress, callback):
        smtpconfig = Thread(target=self._getSmtpConfig, args=(email, progress, callback))
        smtpconfig.start()

    def _getSmtpConfig(self, email, progress, callback):
        url = getUrl(self.ctx, self._configuration.getByName('IspDBUrl'))
        service = 'com.gmail.prrvchr.extensions.OAuth2OOo.OAuth2Service'
        request = createService(self.ctx, service)
        offline = request.isOffLine(url.Server)
        time.sleep(self._progress)
        progress(20)
        time.sleep(self._progress)
        self.InitThread.join()
        progress(40)
        time.sleep(self._progress)
        user, servers = self.DataBase.getSmtpConfig(email)
        print("DataSource._getSmtpConfig() HsqlDB Query user: \n%s\n%s" % (user, servers))
        if len(servers) != 0:
            progress(100, 1)
            print("DataSource._getSmtpConfig() HsqlDB Query")
            callback(user, servers, offline)
        elif offline:
            progress(100, 3)
            print("DataSource._getIspdbConfig() OFFLINE")
            callback(user, (), offline)
        else:
            print("DataSource._getSmtpConfig() IspDB Query")
            progress(60)
            print("DataSource._getSmtpConfig() IspDB Query %s - %s" % (url.Server, url.Complete))
            response = self._getIspdbConfig(request, url.Complete, user.getValue('Domain'))
            if response.IsPresent:
                progress(80)
                config = response.Value
                self.DataBase.setSmtpConfig(config)
                progress(100, 2)
                print("DataSource._getIspdbConfig() OK")
                callback(user, config.getValue('Servers'), offline)
            else:
                progress(100, 4)
                print("DataSource._getIspdbConfig() CANCEL")
                callback(user, (), offline)
   
    def _getIspdbConfig(self, request, url, domain):
        parameter = uno.createUnoStruct('com.sun.star.auth.RestRequestParameter')
        parameter.Method = 'GET'
        parameter.Url = '%s%s' % (url, domain)
        parameter.NoAuth = True
        response = request.getRequest(parameter, DataParser()).execute()
        return response

    def _initDataBase(self):
        try:
            msg = "DataSource for Scheme: loading ... "
            print("DataSource.run() 1 *************************************************************")
            self.DataBase = self._getDataBase()
            print("DataSource.run() 3 *************************************************************")
        except Exception as e:
            msg = "DataSource run(): Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def _getDataBase(self):
        database = DataBase(self.ctx, 'SmtpServer')
        database.addCloseListener(self)
        return database

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
