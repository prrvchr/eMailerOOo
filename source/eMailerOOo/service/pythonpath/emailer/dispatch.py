#!
# -*- coding: utf-8 -*-

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

import uno
import unohelper

from com.sun.star.frame import XNotifyingDispatch

from com.sun.star.frame.DispatchResultState import SUCCESS
from com.sun.star.frame.DispatchResultState import FAILURE

from com.sun.star.uno import Exception as UnoException

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .ispdb import IspdbController

from .merger import MergerController

from .mailer import MailerModel
from .mailer import MailerManager

from .spooler import SpoolerManager
from .spooler import Mailer

from .wizard import Wizard

from .datasource import DataSource

from .unotool import createMessageBox
from .unotool import getPathSettings
from .unotool import getStringResource

from .logger import getLogger
from .logger import RollerHandler

from .helper import checkOAuth2
from .helper import getMailSpooler

from .configuration import g_extension
from .configuration import g_identifier
from .configuration import g_resource
from .configuration import g_basename
from .configuration import g_ispdb_page
from .configuration import g_ispdb_paths
from .configuration import g_merger_page
from .configuration import g_merger_paths
from .configuration import g_defaultlog
from .configuration import g_spoolerlog

import traceback


class Dispatch(unohelper.Base,
               XNotifyingDispatch):
    def __init__(self, ctx, frame):
        self._ctx = ctx
        self._frame = frame
        self._listeners = []

    _datasource = None
    _checked = 0

# XNotifyingDispatch
    def dispatchWithNotification(self, url, arguments, listener):
        state, result = self.dispatch(url, arguments)
        struct = 'com.sun.star.frame.DispatchResultEvent'
        notification = uno.createUnoStruct(struct, self, state, result)
        listener.dispatchFinished(notification)

    def dispatch(self, url, arguments):
        state = FAILURE
        result = ()
        if url.Path == 'ShowIspdb':
            state, result = self._dispatchIspdb(url, arguments)
        else:
            state, result = self._dispatch(url, arguments)
        return state, result

    def addStatusListener(self, listener, url):
        pass

    def removeStatusListener(self, listener, url):
        pass

# Dispatch private methods
    def _dispatch(self, url, arguments):
        state = FAILURE
        result = ()
        # FIXME: We need to check the configuration
        if not self._isChecked(3):
            logger = getLogger(self._ctx, g_defaultlog)
            try:
                datasource = DataSource(self._ctx, self, logger, True)
            except UnoException as e:
                logger.logprb(SEVERE, 'Dispatch', '_dispatch', 1101, url.Main, e.Message)
            else:
                Dispatch._datasource = datasource
                Dispatch._checked |= 3
        # FIXME: Configuration has been checked we can continue
        if self._isChecked(3):
            if url.Path == 'StartSpooler':
                state, result = self._startSpooler()
            elif url.Path == 'StopSpooler':
                state, result = self._stopSpooler()
            elif url.Path == 'ShowSpooler':
                state, result = self._showSpooler()
            elif url.Path == 'ShowMailer':
                state, result = self._showMailer(arguments)
            elif url.Path == 'ShowMerger':
                state, result = self._showMerger()
            elif url.Path == 'GetMail':
                state, result = self._getMail(arguments)
        return state, result

    def _dispatchIspdb(self, url, arguments):
        state = FAILURE
        result = ()
        # FIXME: We need to check the presence of OAuth2OOo extension
        if not self._isChecked(1):
            logger = getLogger(self._ctx, g_defaultlog)
            try:
                checkOAuth2(self._ctx, self, logger, True)
            except UnoException as e:
                logger.logprb(SEVERE, 'Dispatch', '_dispatchIspdb', 1111, url.Main, e.Message)
            else:
                Dispatch._checked |= 1
        if self._isChecked(1):
            state, result = self._showIspdb(arguments)
        return state, result

    #Ispdb methods
    def _showIspdb(self, arguments):
        try:
            state = FAILURE
            email = None
            msg = "Wizard Loading ..."
            sender = ''
            readonly = False
            for argument in arguments:
                if argument.Name == 'Sender':
                    sender = argument.Value
                elif argument.Name == 'ReadOnly':
                    readonly = argument.Value
            parent = self._frame.getContainerWindow()
            wizard = Wizard(self._ctx, g_ispdb_page, True, parent)
            controller = IspdbController(self._ctx, wizard, sender, readonly)
            arguments = (g_ispdb_paths, controller)
            wizard.initialize(arguments)
            msg += " Done ..."
            if wizard.execute() == OK:
                state = SUCCESS
                sender = controller.Model.Sender
                msg +=  " Retrieving SMTP configuration OK..."
            controller.dispose()
            print(msg)
            return state, sender
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.format_exc())
            print(msg)

    #Spooler methods
    def _startSpooler(self):
        getMailSpooler(self._ctx).start()
        return SUCCESS, ()

    def _stopSpooler(self):
        getMailSpooler(self._ctx).terminate()
        return SUCCESS, ()

    def _showSpooler(self):
        try:
            parent = self._frame.getContainerWindow()
            manager = SpoolerManager(self._ctx, self._getDataSource(), parent)
            if manager.execute() == OK:
                manager.saveGrid()
            manager.dispose()
            return SUCCESS, None
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.format_exc())
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
            model = MailerModel(self._ctx, self._getDataSource(), path, close)
            url = model.getDocumentUrl()
            if url is None:
                model.dispose()
            else:
                parent = self._frame.getContainerWindow()
                mailer = MailerManager(self._ctx, model, parent, url)
                if mailer.execute() == OK:
                    state = SUCCESS
                    path = model.getPath()
                mailer.dispose()
            return state, path
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.format_exc())
            print(msg)

    #Merger methods
    def _showMerger(self):
        try:
            msg = "Wizard Loading ..."
            parent = self._frame.getContainerWindow()
            wizard = Wizard(self._ctx, g_merger_page, True, parent)
            controller = MergerController(self._ctx, wizard, self._getDataSource())
            arguments = (g_merger_paths, controller)
            wizard.initialize(arguments)
            msg += " Done ..."
            if wizard.execute() == OK:
                controller.saveGrids()
                msg +=  " Merging SMTP email OK..."
            controller.dispose()
            print(msg)
            return SUCCESS, None
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.format_exc())
            print(msg)

    #Viewer methods
    def _getMail(self, arguments):
        state = FAILURE
        mail = None
        job = None
        for argument in arguments:
            if argument.Name == 'JobId':
                job = argument.Value
        logger = getLogger(self._ctx, g_spoolerlog)
        handler = RollerHandler(self._ctx, logger.Name)
        logger.addRollerHandler(handler)
        if job is None:
            logger.logprb(SEVERE, 'MailSpooler', '_getMail', 1121)
        else:
            logger.logprb(INFO, 'MailSpooler', '_getMail', 1122, job)
            mailer = Mailer(self._ctx, self, self._getDataSource().DataBase, logger)
            try:
                batch, mail = mailer.getMail(job)
            except UnoException as e:
                logger.logprb(SEVERE, 'MailSpooler', '_getMail', 1123, job, e.Message)
            except Exception as e:
                logger.logprb(SEVERE, 'MailSpooler', '_getMail', 1124, job, str(e), traceback.format_exc())
            else:
                logger.logprb(INFO, 'MailSpooler', '_getMail', 1125, job)
                state = SUCCESS
            mailer.dispose()
        logger.removeRollerHandler(handler)
        return state, mail

    # Private methods
    def _isChecked(self, state):
        return Dispatch._checked & state == state

    def _getDataSource(self):
        return Dispatch._datasource

