"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-25 https://prrvchr.github.io                                  ║
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

import unohelper

from com.sun.star.frame.DispatchResultState import FAILURE
from com.sun.star.frame.DispatchResultState import SUCCESS

from com.sun.star.io import XActiveDataControl

from com.sun.star.uno import Exception as UnoException

from com.sun.star.lang import XServiceInfo
from com.sun.star.lang import XComponent

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.ucb.ConnectionMode import ONLINE

from .spooler import Sender

from .unotool import StatusIndicator
from .unotool import TaskEvent

from .unotool import findFrame
from .unotool import getConnectionMode
from .unotool import getDesktop
from .unotool import getNamedValueSet

from .logger import getLogger

from .helper import getMailSpooler

from .configuration import g_spoolerlog
from .configuration import g_spoolerframe
from .configuration import g_dns

from threading import Lock
from time import sleep
import traceback


class MailSend(unohelper.Base,
               XActiveDataControl,
               XServiceInfo,
               XComponent):

    def __init__(self, ctx, implementation, services):
        self._ctx = ctx
        self._implementation = implementation
        self._services = services
        self._lock = Lock()
        self._cls = 'MailSender'
        self._logger = getLogger(ctx, g_spoolerlog)
        self._listeners = []
        self._cancel = TaskEvent()
        self._logger.logprb(INFO, self._cls, '__init__', 601)
        self._thread = None


    # com.sun.star.io.XActiveDataControl
    def addListener(self, listener):
        if listener not in self._listeners:
            self._listeners.append(listener)
        # XXX: Adding listener permit to known the actual state
        if self._hasNoThread():
            listener.closed()
        elif self._isCanceled():
            listener.terminated()
        else:
            listener.started()

    def removeListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

    def start(self):
        with self._lock:
            if self._hasNoThread():
                progress = StatusIndicator(self._ctx, g_spoolerframe)
                self._cancel.clear()
                self._thread = Sender(self._ctx, self._cancel, progress, self._logger, self._listeners)
                self._notifyStarted()

    def terminate(self):
        with self._lock:
            if self._hasThread():
                self._notifyTerminated()
                self._cancel.set()
            else:
                self._notifyClosed()


    # XComponent
    def dispose(self):
        try:
            if self._hasThread():
                self._cancel.set()
                self._thread.join()
        except:
            print("MailSend.dispose() ERROR: %s" % traceback.format_exc())

    def addEventListener(self, listener):
        pass

    def removeEventListener(self, listener):
        pass


    # XServiceInfo
    def supportsService(self, service):
        return service in self._services

    def getImplementationName(self):
        return self._implementation

    def getSupportedServiceNames(self):
        return self._services


    # Private getter methods
    def _isCanceled(self):
        return self._cancel.is_set()

    def _hasThread(self):
        return self._thread and self._thread.is_alive()

    def _hasNoThread(self):
        return self._thread is None or not self._thread.is_alive()

    # Private setter methods
    def _notifyStarted(self):
        for listener in self._listeners:
            listener.started()

    def _notifyClosed(self):
        for listener in self._listeners:
            listener.closed()

    def _notifyTerminated(self):
        for listener in self._listeners:
            listener.terminated()

    def _error(self, e):
        for listener in self._listeners:
            listener.error(e)

