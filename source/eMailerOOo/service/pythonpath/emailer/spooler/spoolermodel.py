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

from com.sun.star.frame.DispatchResultState import SUCCESS

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.util.MeasureUnit import APPFONT

from com.sun.star.view.SelectionType import MULTI

from ..grid import GridManager

from ..unotool import createService
from ..unotool import executeDesktopDispatch
from ..unotool import getConfiguration
from ..unotool import getPathSettings
from ..unotool import getResourceLocation
from ..unotool import getStringResource
from ..unotool import saveWindowPosition

from ..listener import DispatchListener

from ..helper import getMailSpooler

from ..dbtool import getValueFromResult

from ..oauth20 import setParametersArguments

from ..configuration import g_identifier
from ..configuration import g_extension
from ..configuration import g_fetchsize

from string import Template
from urllib import parse
from collections import OrderedDict
import json
import traceback


class SpoolerModel():
    def __init__(self, ctx):
        self._ctx = ctx
        self._closing = False
        self._diposed = False
        self._spooler = None
        self._rowset = None
        self._grid = None
        self._status = 1
        self._indicator = None
        self._path = getPathSettings(ctx).Work
        self._identifiers = ('JobId', )
        self._position = 'SpoolerPosition'
        self._url = getResourceLocation(ctx, g_identifier, 'img')
        self._config = getConfiguration(ctx, g_identifier, True)
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'SpoolerDialog')
        self._resources = {'Title':       'SpoolerDialog.Title',
                           'Button':      'SpoolerDialog.CommandButton1.Label.%s',
                           'Label':       'SpoolerDialog.Label1.Label.%s',
                           'TabTitle':    'SpoolerTab%s.Title',
                           'GridColumns': 'SpoolerTab1.Grid1.Column.%s',
                           'MsgBoxTitle': 'SpoolerMsgBox.Title'}

    @property
    def Path(self):
        return self._path
    @Path.setter
    def Path(self, path):
        self._path = path

    def close(self):
        self._closing = True

# XDispatchResultListener
    def dispatchFinished(self, notification):
        if notification.State == SUCCESS:
            self._path = notification.Result

# SpoolerModel getter methods
    def getSpooler(self):
        return self._spooler

    def getLabelState(self, state):
        return self._getLabelState(state)

    def initToggleState(self):
        state = int(self._spooler.isRunning())
        return state, self._getToggleState(state), self._getLabelState(state)

    def getToggleState(self, state, reset=False):
        if reset:
            self._indicator = None
        print("SpoolerModel.getToggleState() state: %s - label: %s" % (state, self._getLabelState(state)))
        return state, self._getToggleState(state), self._getLabelState(state)

    def hasGridSelectedRows(self):
        return self._grid.hasSelectedRows()

    def getGridSelectedRows(self):
        return self._grid.getSelectedRows()

    def getSelectedColumn(self, column):
        return self._grid.getSelectedColumn(column)

    def resubmitJobs(self, identifier):
        if self._grid.hasSelectedRows():
            jobs = self._grid.getSelectedIdentifiers(identifier)
            self._spooler.resubmitJobs(jobs)

    def isDisposed(self):
        return self._diposed

    def getViewData(self):
        return self._getDialogPosition(), self._getDialogTitle(), self._getTabTitles()

    def getDialogTitle(self):
        return self._getDialogTitle()

    def hasIndicator(self):
        return self._indicator is not None

    def getRowClientInfo(self, selected):
        sent = link = False
        if selected:
            sent = self.getSelectedColumn('State') == 1
            if sent:
                sender = self.getSelectedColumn('Sender')
                client = self._config.getByName('Senders').getByName(sender).getByName('Client')
                if client:
                    link = self._config.getByName('Urls').hasByName(client)
        return selected, sent, link

    def hasDispatch(self):
        return not self._dispatch.isSet() or self._status != 1

    def isSenderStarted(self):
        return self._status == 2

    def endDispatch(self):
        self._dispatch.set()

    def getSelectedJobs(self):
        return self._grid.getSelectedIdentifiers('JobId')

    def startDispatch(self, frame, listener):
        self._dispatch.clear()
        job = self._grid.getSelectedIdentifier('JobId')
        kwargs = {'JobIds': (job, ), 'Frame': frame}
        executeDesktopDispatch(self._ctx, 'emailer:GetMail', listener, **kwargs)

    def getSenderClient(self, sender):
        client = None
        senders = self._config.getByName('Senders')
        if senders.hasByName(sender):
            client = senders.getByName(sender).getByName('Client')
        return client

    def getUrlCommand(self, client, arguments):
        cmd = None
        opt = ''
        urls = self._config.getByName('Urls')
        if urls.hasByName(client):
            url = urls.getByName(client)
            cmd, opt = self._getCommand(url, arguments)
        return cmd, opt

    def getClientCommand(self, arguments):
        return self._getCommand(self._config.getByName('Client'), arguments)

    def getCommandArguments(self):
        sender = self.getSelectedColumn('Sender')
        threadid = self.getSelectedColumn('ThreadId')
        foreignid = self.getSelectedColumn('ForeignId')
        messageid = self.getSelectedColumn('MessageId')
        return {'Sender': sender, 'ThreadId': threadid, 'ForeignId': foreignid, 'MessageId': messageid}

    def _getCommand(self, client, arguments):
        cmd = None
        opt = ''
        parameters = client.getByName('Parameters')
        setParametersArguments(parameters, arguments)
        command = client.getByName('Command')
        if command:
            cmd = Template(command[0]).safe_substitute(arguments)
            option = ' '.join(command[1:]) if len(command) > 1 else ''
            if option:
                opt = Template(option).safe_substitute(arguments)
        return cmd, opt

# SpoolerModel setter methods
    def initSpooler(self, row, spooler):
        print("SpoolerModel.initSpooler() 1")
        self._spooler = getMailSpooler(self._ctx)
        self._spooler.addContentListener(row)
        self._spooler.addSpoolerListener(spooler)
        print("SpoolerModel.initSpooler() 2")

    def startSpooler(self, indicator):
        self._spooler.start(indicator)
        self._indicator = indicator

    def stopSpooler(self):
        self._spooler.stop()
        self._indicator = None

    def addSpoolerIndicator(self, indicator):
        if self._spooler:
            self._spooler.addStatusIndicator(indicator)
            self._indicator = indicator

    def removeSpoolerIndicator(self):
        if self._spooler and self._indicator:
            self._spooler.removeStatusIndicator(self._indicator)
            self._indicator = None

    def isGridInitialized(self):
        return self._grid is not None

    def initGrid(self, rowset, window, listener):
        self._rowset = rowset
        resources = (self._resolver, self._resources.get('GridColumns'))
        self._grid = GridManager(self._ctx, self._url, window, 'Spooler', MULTI, resources)
        self._grid.addSelectionListener(listener)

    def setGridData(self, rowset):
        print("SpoolerModel.setGridData()")
        self._grid.setDataModel(rowset, self._identifiers)

    def deselectAllRows(self):
        self._grid.deselectAllRows()

    def addDocument(self):
        listener = DispatchListener(self)
        kwargs = {'Path': self._path}
        executeDesktopDispatch(self._ctx, 'emailer:ShowMailer', listener, **kwargs)

    def save(self, position):
        saveWindowPosition(self._config, position, self._position)
        self._grid.saveColumnSettings()

    def disposeSpooler(self, row, spooler):
        if self._spooler:
            if row:
                self._spooler.removeContentListener(row)
            if spooler:
                self._spooler.removeSpoolerListener(spooler)
            if self._indicator:
                self._spooler.removeStatusIndicator(self._indicator)
            self._spooler.dispose()
        self._spooler = None
        self._indicator = None

    def dispose(self):
        self._diposed = True
        if self._grid:
            self._grid.dispose()
        self._grid = None
        if self._spooler:
            self._spooler.dispose()
        self._spooler = None

    def removeRows(self, rows):
        jobs = self._getRowsJobs(rows)
        self._spooler.removeJobs(jobs)

    def getMsgBoxTitle(self):
        resource = self._resources.get('MsgBoxTitle')
        return self._resolver.resolveString(resource)

# SpoolerModel private getter methods
    def _getDialogPosition(self):
        return uno.createUnoStruct('com.sun.star.awt.Point', *self._config.getByName(self._position))

    def _getTabTitles(self):
        return self._getTabTitle(1), self._getTabTitle(2), self._getTabTitle(3)

    def _getRowsJobs(self, rows):
        jobs = []
        i = self._rowset.findColumn('JobId')
        for row in rows:
            self._rowset.absolute(row +1)
            jobs.append(getValueFromResult(self._rowset, i))
        return tuple(jobs)

    def _getQueryTable(self):
        table = self._config.getByName('SpoolerTable')
        return table

# SpoolerModel private setter methods
    def _initSpooler(self, sender, window, listener1, listener2, start):
        self._rowset = self._spooler.getContent()
        resources = (self._resolver, self._resources.get('GridColumns'))
        self._grid = GridManager(self._ctx, self._url, window, 'Spooler', MULTI, resources)
        self._grid.addSelectionListener(listener1)
        self._spooler.addContentListener(listener2)
        if start:
            sender.start()

    def _getDialogTitle(self):
        resource = self._resources.get('Title')
        return self._resolver.resolveString(resource)

    def _getTabTitle(self, tabid):
        resource = self._resources.get('TabTitle') % tabid
        return self._resolver.resolveString(resource)

    def _getLabelState(self, state):
        resource = self._resources.get('Label') % state
        return self._resolver.resolveString(resource)

    def _getToggleState(self, state):
        resource = self._resources.get('Button') % state
        return self._resolver.resolveString(resource)

