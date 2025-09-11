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

from com.sun.star.lang import XComponent

from com.sun.star.ui.dialogs import XWizardPage

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .mergerview import MergerView

from ....grid import GridDataListener

from ....dialog import MailManager
from ....dialog import WindowHandler

from ....unotool import getArgumentSet
from ....helper import getMailSpooler

from threading import Condition
import traceback


class MergerManager(MailManager,
                    XWizardPage,
                    XComponent):
    def __init__(self, ctx, wizard, model, pageid, parent):
        super().__init__(ctx, model)
        self._wizard = wizard
        self._pageid = pageid
        self._listeners = []
        self._view = MergerView(ctx, WindowHandler(ctx, self), parent, 2)
        self._view.setSenders(self._model.getSenders())
        self._listener = GridDataListener(self)
        self._model.initPage3(self._listener, self)

# XComponent
    def dispose(self):
        event = uno.createUnoStruct('com.sun.star.lang.EventObject', self)
        for listener in self._listeners:
            listener.disposing(event)
        self._view.getWindow().dispose()
        self._model.getGrid2().removeGridDataListener(self._listener)
        print("MergerManager.dispose() page3")

    def addEventListener(self, listener):
        if listener not in self._listeners:
            self._listeners.append(listener)

    def removeEventListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

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
        if reason == FINISH:
            self.sendDocument()
        return True

    def canAdvance(self):
        return self._canAdvance()

# MergerManager setter methods
    # XXX: This method is called by GridDataListener every time the data in Grid2 is changed
    def dataGridChanged(self):
        recipients, message = self._model.getRecipients()
        self._view.setRecipients(recipients, message)
        self._updateUI()

    def sendDocument(self):
        if self._model.saveDocument():
            subject, attachments = self._getSavedDocumentProperty()
            sender, recipients, filters = self._view.getEmail()
            url, datasource, query, table = self._model.getDocumentInfo()
            spooler = getMailSpooler(self._ctx)
            id = spooler.addMergeJob(sender, subject, url, datasource, query, table, recipients, filters, attachments)
            spooler.dispose()

# MergerManager private setter methods
    def _updateUI(self):
        self._wizard.updateTravelUI()

    def _notifyInit(self, status, result):
        document, message, recipients = result
        if not self._model.isDisposed():
            # TODO: Document can be <None> if a lock or password exists !!!
            # TODO: It would be necessary to test a Handler on the descriptor...
            self._initView(document)
            self._view.setRecipients(getArgumentSet(recipients), message)
            self._updateUI()
        self._model.closeDocument(document)

