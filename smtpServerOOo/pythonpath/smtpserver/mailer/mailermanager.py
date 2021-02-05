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
from unolib import getUrlTransformer
from unolib import parseUrl

from .mailermodel import MailerModel
from .mailerview import MailerView

from smtpserver import logMessage
from smtpserver import getMessage
g_message = 'mailermanager'

import traceback


class MailerManager(unohelper.Base):
    def __init__(self, ctx, datasource, parent, path, recipients=('prrvchr@gmail.com', )):
        self._ctx = ctx
        self._model = MailerModel(ctx, datasource, path)
        self._view = MailerView(ctx, parent, self)
        self._model.getSenders(self.initSenders)
        self._view.setRecipient(recipients)
        print("MailerManager.__init__()")

    @property
    def Model(self):
        return self._model

    def initSenders(self, senders):
        if self._view.isDisposed():
            return
        # Set the Senders ListBox
        self._view.setSenders(senders)

    def initView(self, document):
        if self._view.isDisposed():
            return
        self._model.setDocument(document)
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
        self._view.setAttachAsPdf(int(state))
        # Set the Attachment As PDF CheckBox
        state = self._getDocumentUserProperty('AttachmentAsPdf')
        self._view.setAttachmentAsPdf(int(state))
        # Set the Save Attachments CheckBox and if needed the Attachments ListBox
        state = self._getDocumentUserProperty('SaveAttachments')
        self._view.setSaveAttachments(int(state))
        if state:
            attachments = self._getDocumentAttachemnts('Attachments')
            print("MailerManager.initView() %s" % (attachments,))
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

    def addAttachments(self):
        try:
            resource = self._view.getFilePickerTitleResource()
            attachments = self._model.getAttachments(resource)
            if len(attachments) > 0:
                urls = self._parseAttachments(attachments)
                self._view.addAttachments(urls)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def enableRemoveAttachments(self, enabled):
        self._view.enableRemoveAttachments(enabled)

    def removeAttachments(self):
        self._view.removeAttachments()

    def _parseAttachments(self, attachments):
        urls = []
        transformer = getUrlTransformer(self._ctx)
        pdf = self._view.getAttachmentAsPdf()
        for attachment in attachments:
            url = self._parseAttachment(transformer, attachment, pdf)
            urls.append(url)
        return tuple(urls)

    def _parseAttachment(self, transformer, attachment, pdf):
        url = parseUrl(transformer, attachment)
        if pdf:
            self._addPdfMark(url)
        return transformer.getPresentation(url, False)

    def _addPdfMark(self, url):
        name, extension = self._model.getNameAndExtension(url.Name)
        print("MailerManager._addPdfMark() '%s' - '%s'" % (name, extension))
        if self._model.hasPdfFilter(extension):
            url.Complete = '%s#pdf' % url.Complete

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
