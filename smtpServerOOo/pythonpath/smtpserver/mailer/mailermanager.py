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
    def __init__(self, ctx, manager, datasource, parent, path, recipients=('prrvchr@gmail.com', )):
        self._ctx = ctx
        self._manager = manager
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
        self._canSend()

    def initView(self, document):
        if not self._view.isDisposed():
            #mri = createService(self._ctx, 'mytools.Mri')
            #mri.inspect(document)
            self._model.setUrl(document.URL)
            # Set the Save Subject CheckBox and if needed the Subject TextField
            state = self._getDocumentUserProperty(document, 'SaveSubject')
            self._view.setSaveSubject(int(state))
            if state:
                subject = self._model.getDocumentSubject(document)
                self._view.setSubject(subject)
            # Set the Save Attachments CheckBox and if needed the Attachments ListBox
            state = self._getDocumentUserProperty(document, 'SaveAttachments')
            self._view.setSaveAttachments(int(state))
            if state:
                attachments = self._getDocumentAttachemnts(document, 'Attachments')
                self._view.setAttachments(attachments)
            # Set the Attachment As PDF CheckBox
            state = self._getDocumentUserProperty(document, 'AttachmentAsPdf')
            self._view.setAttachmentAsPdf(int(state))
            # Set the View Document in HTML CommandButton
            self._view.enableButtonViewHtml()
            self._canSend()
        document.close(True)

    def show(self):
        return self._view.execute()

    def dispose(self):
        self._view.dispose()

    def addSender(self, sender):
        self._view.addSender(sender)
        self._canSend()

    def removeSender(self):
        # TODO: button 'RemoveSender' must be deactivated to avoid multiple calls  
        self._view.enableRemoveSender(False)
        sender, position = self._view.getSelectedSender()
        status = self.Model.removeSender(sender)
        if status == 1:
            self._view.removeSender(position)
            self._canSend()

    def editRecipient(self, email, exist):
        enabled = self._validateRecipient(email, exist)
        self._view.enableAddRecipient(enabled)
        self._view.enableRemoveRecipient(exist)

    def addRecipient(self):
        self._view.addRecipient()
        self._canSend()

    def removeRecipient(self):
        self._view.removeRecipient()
        self._canSend()

    def enterRecipient(self, email, exist):
        if self._validateRecipient(email, exist):
            self._view.addToRecipient(email)
            self._canSend()

    def viewHtmlDocument(self):
        document = self._model.getDocument()
        url = self._model.saveDocumentAs(document, 'html')
        document.close(True)
        if url is not None:
            executeShell(self._ctx, url)

    def viewPdfAttachment(self):
        try:
            attachment = self._view.getSelectedAttachment()
            document= self._model.getDocument(attachment)
            url = self._model.saveDocumentAs(document, 'pdf')
            if url is not None:
                executeShell(self._ctx, url)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

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

    def changeAttachments(self, enabled, attachment):
        self._view.enableRemoveAttachments(enabled)
        enabled &= attachment.endswith('#pdf')
        self._view.enableButtonViewPdf(enabled)

    def removeAttachments(self):
        self._view.removeAttachments()

    def changeSender(self, enabled):
        self._view.enableRemoveSender(enabled)

    def changeRecipient(self):
        self._view.enableRemoveRecipient(True)

    def changeSubject(self):
        self._canSend()

    def sendDocument(self):
        try:
            subject = self._view.getSubject()
            attachments = self._view.getAttachments()
            self._saveDocumentProperty(subject, attachments)
            sender = self._view.getSender()
            recipients = self._view.getRecipients()
            url = self._model.getUrl()
            print("MailManager.sendDocument() %s - %s - %s - %s -%s" % (sender, recipients, subject, url, attachments))
            service = 'com.sun.star.mail.MailServiceSpooler'
            spooler = createService(self._ctx, service)
            id = spooler.addJob(sender, subject, url, recipients, attachments)
            print("MailerManager.sendDocument() %s" % id)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def _canSend(self):
        enabled = all((self._view.isSenderValid(),
                       self._view.isRecipientsValid(),
                       self._view.isSubjectValid()))
        self._manager.updateUI(enabled)

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
        if self._model.hasPdfFilter(extension):
            url.Complete = '%s#pdf' % url.Complete

    def _getTitle(self):
        resource = self._view.getTitleRessource()
        title = self._model.getDocumentTitle(resource)
        return title

    def _getDocumentUserProperty(self, document, name):
        resource = self._view.getPropertyResource(name)
        state = self._model.getDocumentUserProperty(document, resource)
        return state

    def _getDocumentAttachemnts(self, document, name):
        resource = self._view.getPropertyResource(name)
        attachments = self._model.getDocumentAttachments(document, resource)
        return attachments

    def _validateRecipient(self, email, exist):
        return all((self.Model.isEmailValid(email), not exist))

    def _saveDocumentProperty(self, subject, attachments):
        try:
            document = self._model.getDocument()
            state = self._view.getSaveSubject()
            self._setDocumentUserProperty(document, 'SaveSubject', state)
            if state:
                self._model.setDocumentSubject(document, subject)
            state = self._view.getSaveAttachments()
            self._setDocumentUserProperty(document, 'SaveAttachments', state)
            if state:
                self._setDocumentAttachments(document, 'Attachments', attachments)
            state = self._view.getAttachmentAsPdf()
            self._setDocumentUserProperty(document, 'AttachmentAsPdf', state)
            document.store()
            document.close(True)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def _setDocumentUserProperty(self, document, name, value):
        resource = self._view.getPropertyResource(name)
        self._model.setDocumentUserProperty(document, resource, value)

    def _setDocumentAttachments(self, document, name, value):
        resource = self._view.getPropertyResource(name)
        self._model.setDocumentAttachments(document, resource, value)
