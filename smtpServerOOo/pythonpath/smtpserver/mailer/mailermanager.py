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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService
from unolib import executeShell

from .mailermodel import MailerModel
from .mailerview import MailerView

from smtpserver import logMessage
from smtpserver import getMessage
g_message = 'mailermanager'

import traceback


class MailerManager(unohelper.Base):
    def __init__(self, ctx, datasource, parent, url, document=None, recipients=('prrvchr@gmail.com', )):
        self._ctx = ctx
        self._model = MailerModel(ctx, datasource)
        self._view = MailerView(ctx, parent, self)
        self._model.getSenders(self.setSenders)
        self._model.getDocument(document, url, self.setDocument)
        self._view.setRecipient(recipients)
        print("MailerManager.__init__()")

    @property
    def Model(self):
        return self._model

    def setSenders(self, senders):
        if self._view.isDisposed():
            return
        # Set the Senders ListBox
        self._view.setSenders(senders)

    def setDocument(self, document):
        if self._view.isDisposed():
            return
        self._model.setDocument(document)
        # Set the dialog Title
        title = self._getTitle()
        self._view.setTitle(title)
        # Set the Save Subject CheckBox and if needed the Subject TextField
        state = self._getDocumentUserProperty('SaveSubject')
        self._view.setSaveSubject(int(state))
        if state:
            subject = self._model.getDocumentSubject()
            self._view.setSubject(subject)
        # Set the Send As HTML / Send As Attachment OptionButton
        state = self._getDocumentUserProperty('SendAsHtml')
        self._view.setSendMode(state)
        # Set the document Name Label
        label = self._getDocumentLabel()
        self._view.setDocumentLabel(label)
        # Set the Save Message CheckBox and if needed the Message TextField
        state = self._getDocumentUserProperty('SaveMessage')
        self._view.setSaveMessage(int(state))
        if state:
            message = self._model.getDocumentDescription()
            self._view.setMessage(message)
        # Set the Attach As PDF CheckBox
        state = self._getDocumentUserProperty('AttachAsPdf')
        self._view.setAttachMode(int(state))
        # Set the Save Attachments CheckBox and if needed the Attachments ListBox
        state = self._getDocumentUserProperty('SaveAttachments')
        self._view.setSaveAttachments(int(state))
        if state:
            attachments = self._getDocumentAttachemnts('Attachments')
            self._view.setAttachments(attachments)
        # Set the View Document in HTML CommandButton
        self._view.enableButtonViewHtml()

    def show(self):
        return self._view.execute()

    def dispose(self):
        self._view.dispose()

    def enableRemoveSender(self, enabled):
        self._view.enableRemoveSender(enabled)

    def addSender(self, sender):
        self._view.addSender(sender)

    def removeSender(self):
        # TODO: button 'RemoveSender' must be deactivated to avoid multiple calls  
        self._view.enableRemoveSender(False)
        sender, position = self._view.getSender()
        status = self.Model.removeSender(sender)
        if status == 1:
            self._view.removeSender(position)

    def editRecipient(self, email, exist):
        enabled = self._validateRecipient(email, exist)
        self._view.enableAddRecipient(enabled)
        self._view.enableRemoveRecipient(exist)

    def addRecipient(self):
        self._view.addToRecipient()

    def changeRecipient(self):
        self._view.enableRemoveRecipient(True)

    def removeRecipient(self):
        self._view.removeRecipient()

    def enterRecipient(self, control, email, exist):
        if self._validateRecipient(email, exist):
            self._view.addRecipient(control, email)

    def sendAsHtml(self):
        self._view.setStep(1)

    def sendAsAttachment(self):
        self._view.setStep(2)

    def viewHtmlDocument(self):
        document = self._model.Document
        url = self._model.saveDocumentAs(document, 'html')
        if url is not None:
            executeShell(self._ctx, url)

    def addAttachment(self):
        resource = self._view.getFilePickerTitleResource()
        documents = self._model.getAttachments(resource)
        print("MailerManager.addAttachment() 1 %s" % (documents, ))
        for document in documents:
            print("MailerManager.addAttachment() 2 %s" % document)

    def removeAttachment(self):
        pass

    def _getTitle(self):
        resource = self._view.getTitleRessource()
        title = self._model.getDocumentTitle(resource)
        return title

    def _getDocumentUserProperty(self, name):
        resource = self._view.getPropertyResource(name)
        state = self._model.getDocumentUserProperty(resource)
        return state

    def _getDocumentLabel(self):
        resource = self._view.getDocumentResource()
        label = self._model.getDocumentLabel(resource)
        return label

    def _getDocumentAttachemnts(self, name):
        resource = self._view.getPropertyResource(name)
        attachments = self._model.getDocumentAttachments(resource)
        return attachments

    def _validateRecipient(self, email, exist):
        return all((self.Model.isEmailValid(email), not exist))
