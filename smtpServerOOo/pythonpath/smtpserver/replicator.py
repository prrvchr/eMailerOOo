#!
# -*- coding: utf_8 -*-

#from __futur__ import absolute_import

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import parseDateTime
from unolib import unparseDateTime

from .dbtools import getDataSource

from .database import DataBase

from .configuration import g_identifier

from .logger import logMessage
from .logger import getMessage

from collections import OrderedDict
from threading import Thread
import traceback
import time


class Replicator(unohelper.Base,
                 Thread):
    def __init__(self, ctx):
        Thread.__init__(self)
        self.ctx = ctx
        #self.datasource = datasource
        self.canceled = False
        #self.sync = sync
        #sync.clear()
        self.error = None
        self.start()

    # XRestReplicator
    def cancel(self):
        self.canceled = True
        self.sync.set()
        self.join()

    def run(self):
        try:
            msg = "Replicator for Scheme: loading ... "
            print("Replicator.run() 1 *************************************************************")
            #logMessage(self.ctx, INFO, "stage 1", 'Replicator', 'run()')
            datasource, url, created = getDataSource(self.ctx, 'SmtpServer', g_identifier, True)
            print("Replicator run() 2")
            self.DataBase = DataBase(self.ctx, datasource)
            if created:
                print("replicator.run() 3")
                self.Error = self.DataBase.createDataBase()
                if self.Error is None:
                    self.DataBase.storeDataBase(url)
            #while not self.canceled:
            print("replicator.run() 4")
            print("replicator.run() 5 *************************************************************")
        except Exception as e:
            msg = "Replicator run(): Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def _synchronize(self):
        pass

    def _syncData(self, timestamp):
        print("Replicator.synchronize() 1")
