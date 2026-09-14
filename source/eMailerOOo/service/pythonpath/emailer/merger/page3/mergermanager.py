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

from com.sun.star.lang import XComponent

from com.sun.star.ui.dialogs import XWizardPage

from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.system.SystemShellExecuteFlags import NO_SYSTEM_ERROR_MESSAGE

from com.sun.star.system import SystemShellExecuteException

from .mergerview import MergerView

from ...mail import MailManager
from ...mail import WindowHandler

from ...grid import GridDataListener

from ...listener import DispatchListener

from ...jobs import Attachment
from ...jobs import Email

from ...runner import Runner
from ...runner import Worker

from ...unotool import Indicator

from ...helper import getMailSpooler

from ...unotool import executeShell

from ...configuration import g_mergerframe
from ...configuration import g_offset

import traceback


class MergerManager(MailManager,
                    XWizardPage,
                    XComponent):
    def __init__(self, ctx, wizard, model, pageid, parent):
        super().__init__(ctx, model)
        self._wizard = wizard
        self._pageid = pageid
        self._listeners = []
        self._view = MergerView(ctx, WindowHandler(self), parent, 2)
        self._view.setSenders(self._model.getSenders())
        self._listener = GridDataListener(self)
        self._indicator = Indicator(ctx=ctx, name=g_mergerframe, offset=g_offset)
        Worker(ctx, self.onInitPage3, self._model.initPage3(), self._listener).start()

    def onInitPage3(self, success, document, recipients, message):
        if not self._model.isClosing() and success:
            # TODO: Document can be <None> if a lock or password exists !!!
            # TODO: It would be necessary to test a Handler on the descriptor...
            self._loadView(document)
            self._view.setRecipients(recipients, message)
            self._updateUI()
        self._model.closeDocument(document)

# XComponent
    def dispose(self):
        self._model.getGrid2().removeGridDataListener(self._listener)
        event = uno.createUnoStruct('com.sun.star.lang.EventObject', self)
        for listener in self._listeners:
            listener.disposing(event)
        self._view.getWindow().dispose()

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

    def getLayout(self):
        return self._view.getLayout()

# MergerManager setter methods
    def viewHTML(self):
        super().viewHTML()
        fltr = self._view.getSelectedRecipient()
        job = Email(self._ctx, self.onViewDocument, *self._model.getEmailData(fltr))
        runner = Runner(self._ctx, job, self._model.getCancel(), self._indicator.createIndicator())
        runner.start()

    def viewPDF(self):
        super().viewPDF()
        fltr = self._view.getSelectedRecipient()
        attachment = self._view.getSelectedAttachment()
        job = Attachment(self._ctx, self.onViewDocument, *self._model.getAttachmentData(fltr, attachment))
        runner = Runner(self._ctx, job, self._model.getCancel(), self._indicator.createIndicator())
        runner.start()

    def onViewDocument(self, success, urls, error):
        if self._model.isClosing():
            return
        self._indicator.resetIndicator()
        self._viewpdf = True
        if success and error is None:
            for url in urls:
                try:
                    executeShell(self._ctx, url, flag=NO_SYSTEM_ERROR_MESSAGE)
                except SystemShellExecuteException:
                    print("MergerManger.onViewDocument() ERROR: %s" % traceback.format_exc())
        self._view.enableViewHtml(self._view.hasSelectedRecipient())
        self._view.enableViewPdf(self._view.hasSelectedPdfAttachment())

    # XXX: This method is called by GridDataListener every time the data in Grid2 is changed
    def dataGridChanged(self):
        if not self._model.isClosing():
            recipients, message = self._model.getRecipients()
            self._view.setRecipients(recipients, message)
            self._updateUI()

    def sendDocument(self):
        if self._model.saveDocument():
            subject, attachments = self._getSavedDocumentProperty()
            sender, recipients, filters = self._view.getEmail()
            spooler = getMailSpooler(self._ctx)
            _ = spooler.addMergeJob(sender, subject, *self._model.getDocumentInfo(recipients, filters, attachments))
            spooler.dispose()

# MergerManager private setter methods
    def _getViewOptions(self, fltr=None):
        return DispatchListener(self), self._model.getTaskEvent(fltr)

    def _updateUI(self):
        self._wizard.updateTravelUI()
