#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-24 https://prrvchr.github.io                                  ║
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

from .mailertool import checkOAuth2
from .mailertool import getDataSource

from .unotool import createMessageBox
from .unotool import getPathSettings
from .unotool import getStringResource

from .logger import getLogger
from .logger import RollerHandler

from .configuration import g_extension
from .configuration import g_identifier
from .configuration import g_resource
from .configuration import g_basename
from .configuration import g_ispdb_page
from .configuration import g_ispdb_paths
from .configuration import g_merger_page
from .configuration import g_merger_paths
from .configuration import g_spoolerlog

import traceback


class MailDispatch(unohelper.Base,
                   XNotifyingDispatch):
    def __init__(self, ctx, parent):
        self._ctx = ctx
        self._parent = parent
        self._listeners = []

    _datasource = None

# XNotifyingDispatch
    def dispatchWithNotification(self, url, arguments, listener):
        state, result = self.dispatch(url, arguments)
        struct = 'com.sun.star.frame.DispatchResultEvent'
        notification = uno.createUnoStruct(struct, self, state, result)
        listener.dispatchFinished(notification)

    def dispatch(self, url, arguments):
        state = FAILURE
        result = None
        if url.Path == 'ispdb':
            oauth2 = None
            # FIXME: We need to check the presence of OAuth2OOo extension
            if not self._isInitialized():
                oauth2 = checkOAuth2(self._ctx, g_extension)
            if oauth2 is not None:
                self._showMsgBox(*oauth2)
            else:
                state, result = self._showIspdb(arguments)
        else:
            # FIXME: We need to check the configuration
            if not self._isInitialized():
                MailDispatch._datasource = getDataSource(self._ctx, g_extension, self._showMsgBox)
            # FIXME: Configuration has been checked we can continue
            if self._isInitialized():
                if url.Path == 'spooler':
                    state, result = self._showSpooler()
                elif url.Path == 'mailer':
                    state, result = self._showMailer(arguments)
                elif url.Path == 'merger':
                    state, result = self._showMerger()
                elif url.Path == 'mail':
                    state, result = self._getMail(arguments)
        return state, result

    def addStatusListener(self, listener, url):
        pass

    def removeStatusListener(self, listener, url):
        pass

# MailDispatch private methods
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
            wizard = Wizard(self._ctx, g_ispdb_page, True, self._parent)
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
    def _showSpooler(self):
        try:
            manager = SpoolerManager(self._ctx, self._getDataSource(), self._parent)
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
                mailer = MailerManager(self._ctx, model, self._parent, url)
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
            wizard = Wizard(self._ctx, g_merger_page, True, self._parent)
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
            logger.logprb(SEVERE, 'MailSpooler', '_getMail()', 1051)
        else:
            logger.logprb(INFO, 'MailSpooler', '_getMail()', 1052, job)
            mailer = Mailer(self._ctx, self, self._getDataSource().DataBase, logger)
            try:
                batch, mail = mailer.getMail(job)
            except UnoException as e:
                logger.logprb(SEVERE, 'MailSpooler', '_getMail()', 1053, job, e.Message)
            except Exception as e:
                logger.logprb(SEVERE, 'MailSpooler', '_getMail()', 1054, job, str(e), traceback.format_exc())
            else:
                logger.logprb(INFO, 'MailSpooler', '_getMail()', 1055, job)
                state = SUCCESS
            mailer.dispose()
        logger.removeRollerHandler(handler)
        return state, mail

    # Private methods
    def _isInitialized(self):
        return MailDispatch._datasource is not None

    def _getDataSource(self):
        return MailDispatch._datasource

    def _showMsgBox(self, method, code, *args):
        resource = getStringResource(self._ctx, g_identifier, g_resource, g_basename)
        message = resource.resolveString(code).format(*args)
        title = resource.resolveString(code +1).format(method)
        msgbox = createMessageBox(self._parent, message, title, 'error', 1)
        msgbox.execute()
        msgbox.dispose()

