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

from com.sun.star.awt.MessageBoxType import WARNINGBOX
from com.sun.star.awt import Point

from com.sun.star.frame.DispatchResultState import FAILURE
from com.sun.star.frame.DispatchResultState import SUCCESS
from com.sun.star.frame import FeatureStateEvent
from com.sun.star.frame import XNotifyingDispatch

from com.sun.star.uno import Exception as UnoException

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.logging.LogLevel import SEVERE

from .spooler import Mailer
from .spooler import Viewer

from .datasource import DataSource

from .unotool import createMessageBox
from .unotool import findFrame
from .unotool import getArgumentSet
from .unotool import getConfiguration
from .unotool import getStringResource

from .logger import getLogger

from .helper import getMailSender

from .ispdb import showIspdbWizard

from .mailer import showMailer

from .merger import showMergerWizard

from .configuration import g_defaultlog
from .configuration import g_identifier
from .configuration import g_ispdb_page
from .configuration import g_ispdb_paths
from .configuration import g_mergerframe
from .configuration import g_merger_page
from .configuration import g_merger_paths
from .configuration import g_spoolerframe
from .configuration import g_spoolerlog
from .configuration import State

import traceback


class Dispatch(unohelper.Base,
               XNotifyingDispatch):
    def __init__(self, ctx, frame):
        self._ctx = ctx
        self._frame = frame
        self._listeners = []
        self._point = Point(*getConfiguration(ctx, g_identifier).getByName('MergerPosition'))
        self._resources = {'Title': 'Dispatch.ErrorBox.Title',
                           'Message': 'Dispatch.ErrorBox.Message'}

    _datasource = None
    _checked = 0

# XNotifyingDispatch
    def dispatchWithNotification(self, url, arguments, notifier):
        self._dispatch(url, arguments, notifier)

    def dispatch(self, url, arguments):
        self._dispatch(url, arguments)

    def addStatusListener(self, listener, url):
        state = FeatureStateEvent()
        state.FeatureURL = url
        state.IsEnabled = True
        #state.State = True
        listener.statusChanged(state)
        self._listeners.append(listener)

    def removeStatusListener(self, listener, url):
        if listener in self._listeners:
            self._listeners.remove(listener)

# Dispatch private methods
    def _dispatch(self, url, arguments, notifier=None):
        try:
            if State.restart:
                self._showMessageBox()
            elif url.Path == 'ShowIspdb':
                showIspdbWizard(self._ctx, dispatch=notifier, **getArgumentSet(arguments))
            else:
                self._dispatchSpooler(url, arguments, notifier)
        except:
            print("Dispatch._dispatch() ERROR: %s" % traceback.format_exc())

    def _dispatchSpooler(self, url, arguments, notifier):
        if url.Path == 'StartSpooler':
            self._startSpooler(notifier)
        elif url.Path == 'StopSpooler':
            self._stopSpooler(notifier)
        elif url.Path == 'ShowSpooler':
            self._showSpooler(notifier)
        elif url.Path == 'ShowMailer':
            self._notifyDispatch(notifier, *showMailer(ctx, **arguments))
        elif url.Path == 'ShowMerger':
            self._showMerger()
        elif url.Path == 'GetMail':
            self._getMail(arguments, notifier)
        elif url.Path == 'GetDocument':
            self._getDocument(arguments, notifier)

    #Spooler methods
    def _startSpooler(self, notifier):
        try:
            frame = findFrame(self._ctx, g_spoolerframe)
            if frame:
                frame.getContainerWindow().toFront()
                getMailSender(self._ctx).start()
                self._notifyDispatch(notifier, SUCCESS)
            else:
                pass
                #manager = SpoolerManager(self._ctx, self._getDataSource(), notifier, True)
        except:
            print("Dispatch._startSpooler() ERROR: %s" % traceback.format_exc())

    def _stopSpooler(self, notifier):
        try:
            frame = findFrame(self._ctx, g_spoolerframe)
            if frame:
                frame.getContainerWindow().toFront()
                getMailSender(self._ctx).terminate()
                state = SUCCESS
            else:
                state = FAILURE
            self._notifyDispatch(notifier, state)
        except:
            print("Dispatch._stopSpooler() ERROR: %s" % traceback.format_exc())

    def _showSpooler(self, notifier):
        try:
            frame = findFrame(self._ctx, g_spoolerframe)
            if frame:
                frame.getContainerWindow().toFront()
            else:
                manager = SpoolerManager(self._ctx, self._getDataSource(), notifier)
        except:
            print("Dispatch._showSpooler() ERROR: %s" % traceback.format_exc())

    #Merger methods
    def _showMerger(self):
        try:
            frame = findFrame(self._ctx, g_mergerframe)
            if frame:
                frame.getContainerWindow().toFront()
            else:
                print("dispatch._showMerger() 1")
                document = self._frame.getController().getModel()
                if document:
                    if document.hasLocation():
                        print("dispatch._showMerger() 2")
                        showMergerWizard(self._ctx, document, point=self._point)
                        print("dispatch._showMerger() 3")
                    else:
                        logger = getLogger(self._ctx, g_defaultlog)
                        title = logger.resolveString(1131)
                        message = logger.resolveString(1132, document.Title)
                        dialog = createMessageBox(self._ctx, title, message, WARNINGBOX)
                        dialog.execute()
                        dialog.dispose()
        except:
            print("Dispatch._showMerger() ERROR: %s" % traceback.format_exc())

    #Viewer methods
    def _getMail(self, arguments, notifier):
        cls = 'Mailer'
        mtd = '_getMail'
        jobs = ()
        event = frame = None
        logger = getLogger(self._ctx, g_spoolerlog)
        for argument in arguments:
            if argument.Name == 'TaskEvent':
                event = argument.Value
            elif argument.Name == 'JobIds':
                jobs = argument.Value
            elif argument.Name == 'Frame':
                frame = argument.Value
        if not jobs:
            logger.logprb(SEVERE, cls, mtd, 1121)
        else:
            try:
                mailer = Mailer(self._ctx, event, frame, logger, jobs, notifier, self)
            except UnoException as e:
                logger.logprb(SEVERE, cls, mtd, 1122, jobslist, e.Message)
            except Exception as e:
                logger.logprb(SEVERE, cls, mtd, 1123, jobslist, str(e), traceback.format_exc())

    def _getDocument(self, arguments, notifier):
        cls = 'Viewer'
        mtd = '_getDocument'
        event = frame = connection = result = datasource = None
        table = url = merge = filter = selection = None
        for argument in arguments:
            if argument.Name == 'TaskEvent':
                event = argument.Value
            elif argument.Name == 'Frame':
                frame = argument.Value
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
            viewer = Viewer(self._ctx, event, frame, connection, result,
                            datasource, table, url, merge, filter, selection, notifier, self)
        except:
            print("Dispatch._getDocument() ERROR: %s" % traceback.format_exc())

    # Private methods

    # Show MessageBox Error
    def _showMessageBox(self):
        resolver = getStringResource(self._ctx, g_identifier, 'dialogs', 'MessageBox')
        title = resolver.resolveString(self._resources.get('Title'))
        message = resolver.resolveString(self._resources.get('Message'))
        dialog = createMessageBox(self._ctx, title, message)
        dialog.execute()
        dialog.dispose()

    def _getDataSource(self):
        return Dispatch._datasource

    def _notifyDispatch(self, notifier, state, result=None):
        if notifier:
            struct = 'com.sun.star.frame.DispatchResultEvent'
            notification = uno.createUnoStruct(struct, self, state, result)
            notifier.dispatchFinished(notification)

