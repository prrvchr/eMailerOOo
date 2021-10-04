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

from com.sun.star.frame import XNotifyingDispatch

from com.sun.star.frame.DispatchResultState import SUCCESS
from com.sun.star.frame.DispatchResultState import FAILURE

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpmailer import DataSource
from smtpmailer import Wizard

from smtpmailer import getMessage
from smtpmailer import getPathSettings
from smtpmailer import logMessage
from smtpmailer import g_extension
from smtpmailer import g_identifier
from smtpmailer import g_ispdb_page
from smtpmailer import g_ispdb_paths
from smtpmailer import g_merger_page
from smtpmailer import g_merger_paths

from .ispdb import IspdbController
from .merger import MergerController
from .sender import SenderModel
from .sender import SenderManager
from .spooler import SpoolerManager

import traceback


class SmtpDispatch(unohelper.Base,
                   XNotifyingDispatch):
    def __init__(self, ctx, parent):
        self._ctx = ctx
        self._parent = parent
        self._listeners = []

    _datasource = None

    @property
    def DataSource(self):
        return SmtpDispatch._datasource

# XNotifyingDispatch
    def dispatchWithNotification(self, url, arguments, listener):
        state, result = self.dispatch(url, arguments)
        struct = 'com.sun.star.frame.DispatchResultEvent'
        notification = uno.createUnoStruct(struct, self, state, result)
        listener.dispatchFinished(notification)

    def dispatch(self, url, arguments):
        if self.DataSource is None:
            SmtpDispatch._datasource = DataSource(self._ctx)
        state = SUCCESS
        result = None
        if url.Path == 'ispdb':
            state, result = self._showIspdb(arguments)
        elif url.Path == 'spooler':
            self._showSpooler()
        elif url.Path == 'mailer':
            state, result = self._showMailer(arguments)
        elif url.Path == 'merger':
            self._showMerger()
        return state, result

    def addStatusListener(self, listener, url):
        pass

    def removeStatusListener(self, listener, url):
        pass

# SmtpDispatch private methods
    #Ispdb methods
    def _showIspdb(self, arguments):
        try:
            state = FAILURE
            email = None
            msg = "Wizard Loading ..."
            close = True
            for argument in arguments:
                if argument.Name == 'Close':
                    close = argument.Value
            wizard = Wizard(self._ctx, g_ispdb_page, True, self._parent)
            controller = IspdbController(self._ctx, wizard, self.DataSource, close)
            arguments = (g_ispdb_paths, controller)
            wizard.initialize(arguments)
            msg += " Done ..."
            if wizard.execute() == OK:
                state = SUCCESS
                email = controller.Model.Email
                msg +=  " Retrieving SMTP configuration OK..."
            controller.dispose()
            print(msg)
            logMessage(self._ctx, INFO, msg, 'SmtpDispatch', '_showIspdb()')
            return state, email
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    #Spooler methods
    def _showSpooler(self):
        try:
            manager = SpoolerManager(self._ctx, self.DataSource, self._parent)
            if manager.execute() == OK:
                manager.saveGrid()
            manager.dispose()
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    #Mailer methods
    def _showMailer(self, arguments):
        try:
            state = FAILURE
            path = None
            close = True
            for argument in arguments:
                if argument.Name == 'Path':
                    path = argument.Value
                elif argument.Name == 'Close':
                    close = argument.Value
            if path is None:
                path = getPathSettings(self._ctx).Work
            model = SenderModel(self._ctx, self.DataSource, path, close)
            url = model.getDocumentUrl()
            if url is not None:
                sender = SenderManager(self._ctx, model, self._parent, url)
                if sender.execute() == OK:
                    state = SUCCESS
                    path = sender.Mailer.Model.Path
                sender.dispose()
            return state, path
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    #Merger methods
    def _showMerger(self):
        try:
            msg = "Wizard Loading ..."
            wizard = Wizard(self._ctx, g_merger_page, True, self._parent)
            controller = MergerController(self._ctx, wizard, self.DataSource)
            arguments = (g_merger_paths, controller)
            wizard.initialize(arguments)
            msg += " Done ..."
            if wizard.execute() == OK:
                controller.saveGrids()
                msg +=  " Merging SMTP email OK..."
            controller.dispose()
            print(msg)
            logMessage(self._ctx, INFO, msg, 'SmtpDispatch', '_showMerger()')
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)
