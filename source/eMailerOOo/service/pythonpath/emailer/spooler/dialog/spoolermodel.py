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

from ...grid import GridManager

from ...unotool import TaskEvent

from ...unotool import createService
from ...unotool import executeDesktopDispatch
from ...unotool import getCallBack
from ...unotool import getConfiguration
from ...unotool import getPathSettings
from ...unotool import getResourceLocation
from ...unotool import getStringResource
from ...unotool import saveTopWindowPosition

from ...listener import DispatchListener

from ...helper import getMailSpooler

from ...dbtool import getValueFromResult

from ...oauth20 import setParametersArguments

from ...configuration import g_identifier
from ...configuration import g_extension
from ...configuration import g_fetchsize

from string import Template
from urllib import parse
from collections import OrderedDict
from threading import Thread
import json
import traceback


class SpoolerModel(unohelper.Base):
    def __init__(self, ctx, datasource):
        self._ctx = ctx
        self._diposed = False
        self._path = getPathSettings(ctx).Work
        self._datasource = datasource
        self._spooler = getMailSpooler(ctx)
        self._rowset = None
        self._grid = None
        self._status = 1
        self._dispatch = TaskEvent(True)
        self._identifiers = ('JobId', )
        self._callback = getCallBack(ctx)
        self._url = getResourceLocation(ctx, g_identifier, 'img')
        self._config = getConfiguration(ctx, g_identifier, True)
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'SpoolerDialog')
        self._resources = {'Title':       'SpoolerDialog.Title',
                           'State':       'SpoolerDialog.Label2.Label.%s',
                           'TabTitle':    'SpoolerTab%s.Title',
                           'GridColumns': 'SpoolerTab1.Grid1.Column.%s'}

    @property
    def _dataSource(self):
        return self._datasource

    @property
    def Path(self):
        return self._path
    @Path.setter
    def Path(self, path):
        self._path = path

# XDispatchResultListener
    def dispatchFinished(self, notification):
        if notification.State == SUCCESS:
            self._path = notification.Result

# SpoolerModel getter methods
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

    def getDialogTitles(self):
        return self._getDialogTitle(), self._getTabTitle(1), self._getTabTitle(2), self._getTabTitle(3)

    def getDialogPosition(self):
        return uno.createUnoStruct('com.sun.star.awt.Point', *self._config.getByName('SpoolerPosition'))

    def getRowClientInfo(self):
        link = False
        sent = self.getSelectedColumn('State') == 1
        if sent:
            sender = self.getSelectedColumn('Sender')
            client = self._config.getByName('Senders').getByName(sender).getByName('Client')
            if client:
                link = self._config.getByName('Urls').hasByName(client)
        return sent, link

    def hasDispatch(self):
        return not self._dispatch.isSet() or self._status != 1

    def isSenderStarted(self):
        return self._status == 2

    def endDispatch(self):
        self._dispatch.set()

    def startDispatch(self, listener):
        self._dispatch.clear()
        job = self._grid.getSelectedIdentifier('JobId')
        kwargs = {'TaskEvent': self._dispatch, 'JobIds': (job, )}
        executeDesktopDispatch(self._ctx, 'emailer:GetMail', listener, **kwargs)

    def setSpoolerStatus(self, status):
        self._status = status
        resource = self._resources.get('State') % status
        label = self._resolver.resolveString(resource)
        return label, int(status == 2), bool(status)

    def getSpoolerStatus(self):
        return self._status

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
    def initSpooler(self, *args):
        Thread(target=self._initSpooler, args=args).start()

    def setGridData(self, rowset):
        self._grid.setDataModel(rowset, self._identifiers)

    def deselectAllRows(self):
        self._grid.deselectAllRows()

    def addDocument(self):
        listener = DispatchListener(self)
        kwargs = {'Path': self._path, 'Close': False}
        executeDesktopDispatch(self._ctx, 'emailer:ShowMailer', listener, **kwargs)

    def save(self, position):
        saveTopWindowPosition(self._config, position, 'SpoolerPosition')
        self._grid.saveColumnSettings()

    def dispose(self):
        self._diposed = True
        if self._grid:
            self._grid.dispose()
        self._datasource.dispose()
        self._spooler.dispose()

    def removeRows(self, rows):
        jobs = self._getRowsJobs(rows)
        self._spooler.removeJobs(jobs)

# SpoolerModel private getter methods
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
    def _initSpooler(self, window, listener1, listener2, caller):
        print("SpoolerModel._initSpooler() wait for database")
        self._rowset = self._spooler.getContent()
        print("SpoolerModel._initSpooler() finish waiting for database")
        resources = (self._resolver, self._resources.get('GridColumns'))
        quote = self._datasource.IdentifierQuoteString
        self._grid = GridManager(self._ctx, self._url, window, quote, 'Spooler', MULTI, resources)
        self._grid.addSelectionListener(listener1)
        self._callback.addCallback(caller, None)
        # TODO: GridColumn and GridModel needs a RowSet already executed!!!
        self._spooler.addContentListener(listener2)

    def _getDialogTitle(self):
        resource = self._resources.get('Title')
        return self._resolver.resolveString(resource)

    def _getTabTitle(self, tabid):
        resource = self._resources.get('TabTitle') % tabid
        return self._resolver.resolveString(resource)

