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

from com.sun.star.ui.dialogs.ExecutableDialogResults import CANCEL
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.system.SystemShellExecuteFlags import NO_SYSTEM_ERROR_MESSAGE

from com.sun.star.system import SystemShellExecuteException

from .mailermodel import MailerModel

from .mailerview import MailerView

from .mailerhandler import DialogHandler

from ..mail import MailManager

from ..mail import WindowHandler

from ..jobs import Attachment
from ..jobs import Email

from ..runner import Runner
from ..runner import Worker

from ..helper import getMailSpooler

from ..unotool import executeShell

import traceback


class MailerManager(MailManager):
    def __init__(self, ctx, url, path, parent=None, **_):
        super().__init__(ctx, MailerModel(ctx, path))
        self._view = MailerView(ctx, DialogHandler(self), WindowHandler(self), parent, 1)
        self._view.setSenders(self._model.getSenders())
        Worker(ctx, self.onInit, self._model.getInitSteps(), url).start()

    def onInit(self, success, document, title):
        if not self._model.isClosing() and success:
            self._loadView(document)
            self._view.setTitle(title)
            self._view.enableButtonCancel()
        if success:
            self._model.closeDocument(document)

# MailerManager getter methods
    def getPath(self):
        return self._model.getPath()

# MailerManager setter methods
    def addRecipient(self):
        email = self._view.getRecipient()
        if self._model.isEmailValid(email):
            self._view.addRecipient(email)
            self._updateUI()

    def changeRecipient(self):
        self._view.enableRemoveRecipient(True)

    def editRecipient(self, email, exist):
        self._view.enableAddRecipient(not exist and self._model.isEmailValid(email))
        self._view.enableRemoveRecipient(exist)

    def enterRecipient(self, email):
        if self._model.isEmailValid(email):
            self._view.addRecipient(email)
            self._updateUI()

    def removeRecipient(self):
        self._view.removeRecipient()
        self._updateUI()

    def execute(self):
        status = self._view.execute()
        self._model.dispose()
        self._view.dispose()
        return status

    def viewHTML(self):
        super().viewHTML()
        job = Email(self._ctx, self.onViewDocument, *self._model.getEmailData())
        runner = Runner(self._ctx, job, self._model.getCancel())
        runner.start()

    def viewPDF(self):
        super().viewPDF()
        attachment = self._view.getSelectedAttachment()
        job = Attachment(self._ctx, self.onViewDocument, *self._model.getAttachmentData(attachment))
        runner = Runner(self._ctx, job, self._model.getCancel())
        runner.start()

    def onViewDocument(self, success, urls, error):
        if self._model.isClosing():
            return
        self._viewpdf = True
        if success and error is None:
            for url in urls:
                try:
                    executeShell(self._ctx, url, flag=NO_SYSTEM_ERROR_MESSAGE)
                except SystemShellExecuteException:
                    print("MailerManger.onViewDocument() ERROR: %s" % traceback.format_exc())
        self._view.enableViewHtml(True)
        self._view.enableViewPdf(self._view.hasSelectedPdfAttachment())

    def sendDocument(self):
        try:
            subject, attachments = self._getSavedDocumentProperty()
            sender, recipients = self._view.getEmail()
            url = self._model.getUrl()
            spooler = getMailSpooler(self._ctx)
            _ = spooler.addJob(sender, subject, url, recipients, attachments)
            spooler.dispose()
            self._view.endDialog(OK)
        except Exception:
            print("MailerManager.sendDocument() ERROR: %s" % traceback.format_exc())

    def cancel(self):
        self._view.endDialog(CANCEL)

# MailerManager private setter methods
    def _updateUI(self):
        enabled = self._canAdvance()
        self._view.enableButtonSend(enabled)
