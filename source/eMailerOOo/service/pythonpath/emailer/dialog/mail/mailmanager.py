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

from com.sun.star.awt.MessageBoxType import WARNINGBOX

from com.sun.star.awt import XCallback

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from ...unotool import createMessageBox
from ...unotool import getArgumentSet
from ...unotool import getFileUrl

from threading import Lock
import traceback


class MailManager(unohelper.Base,
                  XCallback):
    def __init__(self, ctx, model):
        self._ctx = ctx
        self._model = model
        self._lock = Lock()
        self._viewpdf = True
        self._disabled = False
        self._background = {True: 16777215, False: 16711680}

    @property
    def Model(self):
        return self._model

    @property
    def View(self):
        return self._view

    # TODO: One shot disabler handler
    def isHandlerEnabled(self):
        enabled = True
        if self._disabled:
            self._disabled = enabled = False
        return enabled
    def _disableHandler(self):
        self._disabled = True

# XCallback
    def notify(self, data):
        self._notify(**getArgumentSet(data))

# XDispatchResultListener
    def dispatchFinished(self, notification):
        state = notification.State
        result = notification.Result
        self._dispatchFinished(result.First, state, result.Second)

    def _dispatchFinished(self, call, state, value):
        raise NotImplementedError('Need to be implemented!')

# MailManager setter methods
    def addRecipient(self):
        # FIXME: noops method called from the WindowHandler
        pass

    def changeRecipient(self):
        raise NotImplementedError('Need to be implemented!')

    def editRecipient(self, email, exist):
        # FIXME: noops method called from the WindowHandler
        pass

    def enterRecipient(self, email):
        # FIXME: noops method called from the WindowHandler
        pass

    def removeRecipient(self):
        # FIXME: noops method called from the WindowHandler
        pass

    def sendDocument(self):
        raise NotImplementedError('Need to be implemented!')

    def removeSender(self):
        # TODO: button 'RemoveSender' must be deactivated to avoid multiple calls  
        self._view.enableRemoveSender(False)
        sender, position = self._view.getSelectedSender()
        if self.Model.removeSender(sender):
            self._view.removeSender(position)
            self._updateUI()

    def _getViewOptions(self, filter=None):
        return (self, )

    def viewHtml(self):
        self._view.enableViewHtml(False)
        self._view.enableViewPdf(False)
        self._viewpdf = False
        selection = self._view.getSelectedRecipients()
        url = self._model.parseUri()
        options = self._getViewOptions('html')
        self._model.viewDocument(selection, url, True, 'html', *options)

    def viewPdf(self):
        self._view.enableViewHtml(False)
        self._view.enableViewPdf(False)
        self._viewpdf = False
        selection = self._view.getSelectedRecipients()
        url, merge, filter = self._model.parseUriFragment(self._view.getSelectedAttachment())
        options = self._getViewOptions('pdf')
        self._model.viewDocument(selection, url, merge, 'pdf', *options)

    def changeAttachments(self, index, selected, item, positions):
        self._view.enableRemoveAttachments(selected)
        enabled = selected and min(positions) > 0
        self._view.enableMoveBefore(enabled)
        enabled = selected and max(positions) < index
        self._view.enableMoveAfter(enabled)
        enabled = selected and item.endswith('pdf')
        self._view.enableViewPdf(enabled and self._viewpdf)

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

    def changeSubject(self):
        self._updateUI()

# MailManager private getter methods
    def _initView(self, document):
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
        self._view.enableViewHtml(True)
        # Set the View Message Label
        message = self._model.getDocumentMessage(document)
        self._view.setMessage(message)
        self._updateUI()

    def _canAdvance(self):
        valid = self._model.isSubjectValid(self._view.getSubject())
        self._view.setSubjectBackground(self._background.get(valid))
        sender = self._view.getSender()
        recipients = self._view.getRecipients()
        return valid and self._model.isSenderValid(sender) and self._model.isRecipientsValid(recipients)

    def _getSavedDocumentProperty(self):
        subject = self._view.getSubject()
        attachments = self._view.getAttachments()
        document = self._model.getDocument(None, False)
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
        self._model.closeDocument(document)
        return subject, attachments

# MailManager private setter methods
    def _updateUI(self):
        raise NotImplementedError('Need to be implemented!')

    def _notify(self, call, **kwargs):
        with self._lock:
            if call == 'init':
                self._notifyInit(**kwargs)
            else:
                self._notifyView(**kwargs)
                self._viewpdf = True
                self._view.enableViewHtml(True)
                self._view.enableViewPdf(self._view.hasSelectedPdfAttachment())

    def _notifyInit(self, **kwargs):
        raise NotImplementedError('Need to be implemented!')

    def _notifyView(self, status, result):
        raise NotImplementedError('Need to be implemented!')

    def _showMessageBox(self, message):
        parent = self._view.getWindow().Peer
        title = self._model.getMsgBoxTitle()
        dialog = createMessageBox(parent, WARNINGBOX, 1, title, message)
        dialog.execute()
        dialog.dispose()

