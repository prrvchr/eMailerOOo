#!
# -*- coding: utf-8 -*-

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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpmailer import executeShell
from smtpmailer import getFileUrl
from smtpmailer import getMessage
from smtpmailer import logMessage

g_message = 'basemanager'

import traceback


class MailManager(unohelper.Base):
    def __init__(self):
        raise NotImplementedError('Need to be implemented!')

    @property
    def Model(self):
        return self._model

    # TODO: One shot disabler handler
    def isHandlerEnabled(self):
        enabled = True
        if self._disabled:
            self._disabled = enabled = False
        return enabled
    def _disableHandler(self):
        self._disabled = True

# MailManager setter methods
    def sendDocument(self):
        raise NotImplementedError('Need to be implemented!')

    def initSenders(self, senders):
        with self._lock:
            if not self._model.isDisposed():
                # Set the Senders ListBox
                self._view.setSenders(senders)
                self._updateUI()

    def initView(self, document):
        with self._lock:
            if not self._model.isDisposed():
                self._model.setUrl(document.URL)
                # Set the Save Subject CheckBox and if needed the Subject TextField
                state = self._model.getDocumentUserProperty(document, 'SaveSubject')
                self._view.setSaveSubject(int(state))
                if state:
                    subject = self._model.getDocumentSubject(document)
                    self._view.setSubject(subject)
                # Set the Save Attachments CheckBox and if needed the Attachments ListBox
                state = self._model.getDocumentUserProperty(document, 'SaveAttachments')
                self._view.setSaveAttachments(int(state))
                if state:
                    attachments = self._model.getDocumentAttachemnts(document, 'Attachments')
                    self._view.setAttachments(attachments)
                # Set the Attachment As PDF CheckBox
                state = self._model.getDocumentUserProperty(document, 'AttachmentAsPdf')
                self._view.setAttachmentAsPdf(int(state))
                # Set the View Document in HTML CommandButton
                self._view.enableViewHtml()
                # Set the View Message Label
                message = self._model.getDocumentMessage(document)
                self._view.setMessage(message)
                self._updateUI()
            self._closeDocument(document)

    def addSender(self, sender):
        self._view.addSender(sender)
        self._updateUI()

    def removeSender(self):
        # TODO: button 'RemoveSender' must be deactivated to avoid multiple calls  
        self._view.enableRemoveSender(False)
        sender, position = self._view.getSelectedSender()
        status = self.Model.removeSender(sender)
        if status == 1:
            self._view.removeSender(position)
            self._updateUI()

    def editRecipient(self, email, exist):
        enabled = self._model.validateRecipient(email, exist)
        self._view.enableAddRecipient(enabled)
        self._view.enableRemoveRecipient(exist)

    def addRecipient(self):
        self._view.addRecipient()
        self._updateUI()

    def removeRecipient(self):
        self._view.removeRecipient()
        self._updateUI()

    def enterRecipient(self, email, exist):
        if self._model.validateRecipient(email, exist):
            self._view.addToRecipient(email)
            self._updateUI()

    def viewHtml(self):
        index = self._view.getCurrentRecipient()
        document = self._model.getDocument()
        if index is not None:
            self._model.mergeDocument(document, index)
        url = self._model.saveDocumentAs(document, 'html')
        self._closeDocument(document)
        if url is not None:
            executeShell(self._ctx, url)

    def viewPdf(self):
        index = self._view.getCurrentRecipient()
        attachment = self._view.getSelectedAttachment()
        document = self._model.getDocument(attachment)
        if index is not None and self._model.hasMergeMark(attachment):
            self._model.mergeDocument(document, index)
        url = self._model.saveDocumentAs(document, 'pdf')
        self._closeDocument(document)
        if url is not None:
            executeShell(self._ctx, url)

    def changeAttachments(self, index, selected, item, positions):
        self._view.enableRemoveAttachments(selected)
        enabled = selected and min(positions) > 0
        self._view.enableMoveBefore(enabled)
        enabled = selected and max(positions) < index
        self._view.enableMoveAfter(enabled)
        enabled = selected and item.endswith('pdf')
        self._view.enableViewPdf(enabled)

    def moveAttachments(self, offset):
        self._view.enableRemoveAttachments(False)
        self._view.enableMoveBefore(False)
        self._view.enableMoveAfter(False)
        self._view.enableViewPdf(False)
        # TODO: We must disable the handler "ChangeAttachments" otherwise it activates twice
        self._disableHandler()
        self._view.moveAttachments(offset)

    def addAttachments(self):
        title = self._model.getFilePickerTitle()
        path = self._model.Path
        urls, path = getFileUrl(self._ctx, title, path, (), True)
        self._model.Path = path
        if urls is not None:
            merge = self._view.getMergeAttachments()
            pdf = self._view.getAttachmentAsPdf()
            attachments = self._model.parseAttachments(urls, merge, pdf)
            self._view.addAttachments(attachments)

    def removeAttachments(self):
        self._view.removeAttachments()

    def changeSender(self, enabled):
        self._view.enableRemoveSender(enabled)

    def changeRecipient(self):
        self._view.enableRemoveRecipient(True)

    def changeSubject(self):
        self._updateUI()

# MailManager private getter methods
    def _canAdvance(self):
        valid = all((self._view.isSenderValid(),
                     self._view.isRecipientsValid(),
                     self._view.isSubjectValid()))
        return valid

    def _getSavedDocumentProperty(self):
        subject = self._view.getSubject()
        attachments = self._view.getAttachments()
        print("MailManager.getAttachments() 1 ******************************** %s" % (attachments, ))
        document = self._model.getDocument()
        state = self._view.getSaveSubject()
        self._model.setDocumentUserProperty(document, 'SaveSubject', state)
        if state:
            self._model.setDocumentSubject(document, subject)
        state = self._view.getSaveAttachments()
        self._model.setDocumentUserProperty(document, 'SaveAttachments', state)
        if state:
            self._model.setDocumentAttachments(document, 'Attachments', attachments)
            state = self._view.getAttachmentAsPdf()
            self._model.setDocumentUserProperty(document, 'AttachmentAsPdf', state)
        document.store()
        self._closeDocument(document)
        return subject, attachments

# MailManager private setter methods
    def _updateUI(self):
        raise NotImplementedError('Need to be implemented!')

    def _closeDocument(self, document):
        raise NotImplementedError('Need to be implemented!')
