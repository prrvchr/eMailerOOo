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

from com.sun.star.ui.dialogs import XWizardPage

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService

from .mergerview import MergerView

from smtpserver import logMessage
from smtpserver import getMessage

import traceback


class MergerManager(unohelper.Base,
                    XWizardPage):
    def __init__(self, ctx, wizard, model, pageid, parent):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._disabled = False
        datasources = self._model.getAvailableDataSources()
        self._view = MergerView(ctx, self, parent, datasources)
        datasource = self._model.getDocumentDataSource()
        if datasource in datasources:
            self._view.setPageStep(1)
            # TODO: We must disable the handler "ChangeDataSource" otherwise it activates twice
            self._disableHandler()
            self._view.selectDataSource(datasource)
        else:
            self._view.enableDatasource(True)

    @property
    def Model(self):
        return self._model

    def isHandlerEnabled(self):
        enabled = True
        if self._disabled:
            self._disabled = enabled = False
        return enabled
    def _disableHandler(self):
        self._disabled = True

# XWizardPage
    @property
    def PageId(self):
        return self._pageid
    @property
    def Window(self):
        return self._view.getWindow()

    def activatePage(self):
        pass

    def commitPage(self, reason):
        return True

    def canAdvance(self):
        return all((self._view.isQuerySelected(),
                    self._view.hasEmail(),
                    self._view.hasIndex()))

# MergerManager setter methods
    # DataSource setter methods
    def changeDataSource(self, datasource):
        self._view.enablePage(False)
        self._view.enableButton(False)
        self._view.setPageStep(1)
        self._model.setDataSource(datasource, self.progress, self.setDataSource)

    def progress(self, value):
        self._view.updateProgress(value)

    def setDataSource(self, step, queries, tables, label1, label2, msg):
        if step == 2:
            self._view.setMessageText(msg)
        elif step == 3:
            # TODO: We must disable the handler "EditQuery" otherwise it activates twice
            self._disableHandler()
            self._view.initQuery(queries)
            # TODO: We must disable the handler "ChangeTable" otherwise it activates twice
            self._disableHandler()
            self._view.initTables(tables)
            self._view.setEmailLabel(label1)
            self._view.setIndexLabel(label2)
            self._view.enablePage(True)
        self._view.setPageStep(step)
        self._wizard.updateTravelUI()

    def newDataSource(self):
        frame = self._model.getCurrentDocument().CurrentController.Frame
        service = 'com.sun.star.frame.DispatchHelper'
        dispatcher = createService(self._ctx, service)
        dispatcher.executeDispatch(frame, '.uno:AutoPilotAddressDataSource', '', 0, ())
        #elif tag == 'AddressBook':
        #    dispatcher.executeDispatch(frame, '.uno:AddressBookSource', '', 0, ())
        # TODO: Update the list of data sources and keep the selection if possible
        datasource = self._view.getDataSource()
        datasources = self._model.getAvailableDataSources()
        self._view.initDataSource(datasources)
        if datasource in datasources:
            # TODO: We must disable the handler "ChangeDataSource" otherwise it activates twice
            self._disableHandler()
            self._view.selectDataSource(datasource)

    # Table setter methods
    def changeTables(self, table):
        columns, emails = self._model.setTable(table)
        # TODO: We must disable the handler "ChangeColumn" otherwise it activates twice
        self._disableHandler()
        self._view.initColumns(columns)
        self._view.setEmail(emails)

    # Columns setter methods
    def changeColumns(self, column):
        emails = self._view.getEmails()
        enabled = column not in emails
        self._view.enableAddEmail(enabled)
        indexes = self._view.getIndexes()
        enabled = all((self._view.isQuerySelected(),
                       column not in indexes))
        self._view.enableAddIndex(enabled)

    # Query setter methods
    def editQuery(self, query, exist):
        enabled = self._model.validateQuery(query, exist)
        indexes = self._model.setQuery(query) if exist else ()
        self._setQuery(indexes, exist)
        self._view.enableAddQuery(enabled)
        self._wizard.updateTravelUI()

    def changeQuery(self, query):
        indexes = self._model.setQuery(query)
        self._setQuery(indexes, True)
        self._wizard.updateTravelUI()

    def enterQuery(self, query, exist):
        if self._model.validateQuery(query, exist):
            self._addQuery(query)
            self._wizard.updateTravelUI()

    def addQuery(self):
        query = self._view.getQuery()
        self._addQuery(query)
        self._wizard.updateTravelUI()

    def removeQuery(self):
        query = self._view.getQuery()
        self._model.removeQuery(query)
        self._view.removeQuery(query)
        self._wizard.updateTravelUI()

    # Email column setter methods
    def changeEmail(self, imax, position):
        self._view.enableRemoveEmail(position != -1)
        self._view.enableBefore(position > 0)
        self._view.enableAfter(-1 < position < imax)

    def addEmail(self):
        self._view.enableAddEmail(False)
        self._view.enableRemoveEmail(False)
        self._view.enableBefore(False)
        self._view.enableAfter(False)
        table = self._view.getTable()
        email = self._view.getColumn()
        emails = self._model.addEmail(table, email)
        self._view.setEmail(emails)
        self._wizard.updateTravelUI()

    def removeEmail(self):
        self._view.enableAddEmail(False)
        self._view.enableRemoveEmail(False)
        self._view.enableBefore(False)
        self._view.enableAfter(False)
        table = self._view.getTable()
        email = self._view.getEmail()
        emails = self._model.removeEmail(table, email)
        self._view.setEmail(emails)
        column = self._view.getColumn()
        if column not in emails:
            self._view.enableAddEmail(True)
        self._wizard.updateTravelUI()

    def moveBefore(self):
        self._view.enableRemoveEmail(False)
        self._view.enableBefore(False)
        self._view.enableAfter(False)
        table = self._view.getTable()
        email = self._view.getEmail()
        index = self._view.getEmailPosition() -1
        emails = self._model.moveEmail(table, email, index)
        self._view.setEmail(emails, index)

    def moveAfter(self):
        self._view.enableRemoveEmail(False)
        self._view.enableBefore(False)
        self._view.enableAfter(False)
        table = self._view.getTable()
        email = self._view.getEmail()
        index = self._view.getEmailPosition() +1
        emails = self._model.moveEmail(table, email, index)
        self._view.setEmail(emails, index)

    # Index column setter methods
    def changeIndex(self, enabled):
        self._view.enableRemoveIndex(enabled)

    def addIndex(self):
        self._view.enableAddIndex(False)
        self._view.enableRemoveIndex(False)
        index = self._view.getColumn()
        indexes = self._model.addIndex(index)
        self._view.setIndexes(indexes)
        self._wizard.updateTravelUI()

    def removeIndex(self):
        self._view.enableRemoveIndex(False)
        self._view.enableAddIndex(False)
        index = self._view.getIndex()
        indexes = self._model.removeIndex(index)
        self._view.setIndexes(indexes)
        column = self._view.getColumn()
        if column not in indexes:
            self._view.enableAddIndex(True)
        self._wizard.updateTravelUI()

# MergerManager private setter methods
    # Query private setter methods
    def _setQuery(self, indexes, exist):
        self._view.setIndexes(indexes)
        column = self._view.getColumn()
        enabled = column not in indexes
        self._view.enableAddIndex(enabled)
        self._view.enableRemoveQuery(exist)

    def _addQuery(self, query):
        table = self._view.getTable()
        self._model.addQuery(table, query)
        self._view.addQuery(query)
        self._view.enableAddIndex(False)
        self._view.enableRemoveQuery(False)
