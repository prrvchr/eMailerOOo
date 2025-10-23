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

from com.sun.star.awt import Point
from com.sun.star.awt.WindowClass import MODALTOP

from com.sun.star.frame.DispatchResultState import FAILURE
from com.sun.star.frame.DispatchResultState import SUCCESS
from com.sun.star.frame.FrameSearchFlag import GLOBAL
from com.sun.star.frame import XNotifyingDispatch

from com.sun.star.uno import Exception as UnoException

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .wizard import IspdbController
from .wizard import MergerController

from .dialog import MailerModel
from .dialog import MailerManager

from .spooler import Mailer
from .spooler import SpoolerManager
from .spooler import Viewer

from .wizard import Wizard

from .datasource import DataSource

from .unotool import StatusIndicator

from .unotool import createMessageBox
from .unotool import getConfiguration
from .unotool import getDesktop
from .unotool import getPathSettings
from .unotool import getStringResource

from .logger import getLogger
from .logger import RollerHandler

from .helper import checkOAuth2
from .helper import getMailSender

from .configuration import g_extension
from .configuration import g_identifier
from .configuration import g_resource
from .configuration import g_basename
from .configuration import g_ispdb_page
from .configuration import g_ispdb_paths
from .configuration import g_mergerframe
from .configuration import g_merger_page
from .configuration import g_merger_paths
from .configuration import g_spoolerframe
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
    def dispatchWithNotification(self, url, arguments, notifier):
        self._dispatch(url, arguments, notifier)
        print("Dispatch.dispatchWithNotification () finished...")

    def dispatch(self, url, arguments):
        self._dispatch(url, arguments)
        print("Dispatch.dispatch () finished...")

    def addStatusListener(self, listener, url):
        pass

    def removeStatusListener(self, listener, url):
        pass

# Dispatch private methods
    def _dispatch(self, url, arguments, notifier=None):
        try:
            if url.Path == 'ShowIspdb':
                self._dispatchIspdb(url, arguments, notifier)
            else:
                self._dispatchSpooler(url, arguments, notifier)
        except Exception as e:
            msg = "Dispatch._dispatch() Error: %s - %s" % (e, traceback.format_exc())
            print(msg)

    def _dispatchIspdb(self, url, arguments, notifier):
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
            self._showIspdb(arguments, notifier)

    def _dispatchSpooler(self, url, arguments, notifier):
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
                self._showSpooler(notifier)
            elif url.Path == 'ShowMailer':
                self._notifyDispatch(notifier, *self._showMailer(arguments))
            elif url.Path == 'ShowMerger':
                self._showMerger()
            elif url.Path == 'GetMail':
                self._getMail(arguments, notifier)
            elif url.Path == 'GetDocument':
                self._getDocument(arguments, notifier)

    #Ispdb methods
    def _showIspdb(self, arguments, notifier):
        try:
            print("Dispatch._showIspdb() 1")
            email = None
            msg = "Wizard Loading ..."
            sender = ''
            parent = None
            readonly = False
            for argument in arguments:
                if argument.Name == 'Sender':
                    sender = argument.Value
                elif argument.Name == 'ParentWindow':
                    parent = argument.Value
                elif argument.Name == 'ReadOnly':
                    readonly = argument.Value
            if parent is None:
                parent = self._frame.getContainerWindow().getToolkit().getActiveTopWindow()
            wizard = Wizard(self._ctx, g_ispdb_page, True, parent)
            controller = IspdbController(self._ctx, wizard, sender, readonly, notifier)
            arguments = (g_ispdb_paths, controller)
            wizard.initialize(arguments)
            msg += " Done ..."
            if wizard.execute() == OK:
                msg +=  " Retrieving SMTP configuration OK..."
            wizard.dispose()
            print(msg)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.format_exc())
            print(msg)

    #Spooler methods
    def _startSpooler(self):
        try:
            frame = getDesktop(self._ctx).findFrame(g_spoolerframe, GLOBAL)
            if frame:
                frame.getContainerWindow().toFront()
            else:
                manager = SpoolerManager(self._ctx, self._getDataSource())
            getMailSender(self._ctx).start()
            return SUCCESS, ()
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.format_exc())
            print(msg)

    def _stopSpooler(self):
        try:
            frame = getDesktop(self._ctx).findFrame(g_spoolerframe, GLOBAL)
            if frame:
                frame.getContainerWindow().toFront()
                getMailSender(self._ctx).terminate()
                state = SUCCESS
            else:
                state = FAILURE
            return state, ()
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.format_exc())
            print(msg)

    def _showSpooler(self, notifier):
        try:
            print("Dispatch._showSpooler() 1")
            frame = getDesktop(self._ctx).findFrame(g_spoolerframe, GLOBAL)
            print("Dispatch._showSpooler() 2")
            if frame:
                print("Dispatch._showSpooler() 3")
                frame.getContainerWindow().toFront()
            else:
                print("Dispatch._showSpooler() 4")
                manager = SpoolerManager(self._ctx, self._getDataSource(), notifier)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.format_exc())
            print(msg)

    #Mail methods
    def _showMailer(self, arguments):
        try:
            state = FAILURE
            path = None
            for argument in arguments:
                if argument.Name == 'Path':
                    path = argument.Value
            if path is None:
                path = getPathSettings(self._ctx).Work
            model = MailerModel(self._ctx, path)
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
            frame = getDesktop(self._ctx).findFrame(g_mergerframe, GLOBAL)
            if frame:
                frame.getContainerWindow().toFront()
            else:
                msg = "Wizard Loading ..."
                document = self._frame.getController().getModel()
                if document.hasLocation():
                    config = getConfiguration(self._ctx, g_identifier)
                    if config.hasByName('MergerPosition'):
                        point = Point(*config.getByName('MergerPosition'))
                    else:
                        point = None
                    config.dispose()
                    print("Dispatch._showMerger() 4 Point X: %s - Y: %s" % (point.X, point.Y))
                    print("Dispatch._showMerger() 5 document.URL: %s" % document.URL)
                    wizard = Wizard(self._ctx, g_merger_page, True, None, g_mergerframe, point)
                    controller = MergerController(self._ctx, wizard, document)
                    arguments = (g_merger_paths, controller)
                    wizard.initialize(arguments)
                    msg += " Done ..."
                    print("Dispatch._showMerger() 6 document.URL: %s" % document.URL)
                    wizard.execute()
                else:
                    logger = getLogger(self._ctx, g_defaultlog)
                    box = uno.Enum('com.sun.star.awt.MessageBoxType', 'WARNINGBOX')
                    title = logger.resolveString(1131)
                    message = logger.resolveString(1132, document.Title)
                    dialog = createMessageBox(parent, box, 1, title, message)
                    dialog.execute()
                    dialog.dispose()
            print(msg)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.format_exc())
            print(msg)

    #Viewer methods
    def _getMail(self, arguments, notifier):
        cls = 'MailSpooler'
        mtd = '_getMail'
        jobs = ()
        event = None
        logger = getLogger(self._ctx, g_spoolerlog)
        for argument in arguments:
            if argument.Name == 'TaskEvent':
                event = argument.Value
            elif argument.Name == 'JobIds':
                jobs = argument.Value
        if not jobs:
            logger.logprb(SEVERE, cls, mtd, 1121)
        else:
            try:
                progress = StatusIndicator(self._ctx, g_spoolerframe)
                mailer = Mailer(self._ctx, event, progress, logger, jobs, notifier, self)
            except UnoException as e:
                logger.logprb(SEVERE, cls, mtd, 1122, jobslist, e.Message)
            except Exception as e:
                logger.logprb(SEVERE, cls, mtd, 1123, jobslist, str(e), traceback.format_exc())

    def _getDocument(self, arguments, notifier):
        cls = 'Merger'
        mtd = '_getDocument'
        event = connection = result = datasource = table = url = merge = filter = selection = None
        logger = getLogger(self._ctx, g_spoolerlog)
        for argument in arguments:
            if argument.Name == 'TaskEvent':
                event = argument.Value
            elif argument.Name == 'Connection':
                connection = argument.Value
            elif argument.Name == 'ResultSet':
                result = argument.Value
            elif argument.Name == 'DataSource':
                datasource = argument.Value
            elif argument.Name == 'Table':
                table = argument.Value
            elif argument.Name == 'Url':
                url = argument.Value
            elif argument.Name == 'Merge':
                merge = argument.Value
            elif argument.Name == 'Filter':
                filter = argument.Value
            elif argument.Name == 'Selection':
                selection = argument.Value
        try:
            progress = StatusIndicator(self._ctx, g_mergerframe)
            viewer = Viewer(self._ctx, event, progress, connection, result,
                            datasource, table, url, merge, filter, selection, notifier, self)
        except Exception as e:
            print("Dispatch._getDocument() ERROR: %s" % traceback.format_exc())


    # Private methods
    def _isChecked(self, state):
        return Dispatch._checked & state == state

    def _getDataSource(self):
        return Dispatch._datasource

    def _notifyDispatch(self, notifier, state, result):
        if notifier:
            struct = 'com.sun.star.frame.DispatchResultEvent'
            notification = uno.createUnoStruct(struct, self, state, result)
            notifier.dispatchFinished(notification)

