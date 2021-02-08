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

from unolib import getDialog
from unolib import getContainerWindow

from .mailerhandler import WindowHandler

from smtpserver import g_extension

import traceback


class MailerView(unohelper.Base):
    def __init__(self, ctx, parent, manager):
        handler = WindowHandler(ctx, manager)
        self._window = getContainerWindow(ctx, parent, handler, g_extension, 'MailerWindow')
        self._window.setVisible(True)

# MailerView setter methods
    def setRecipient(self, recipients):
        if len(recipients) > 0:
            control = self._getRecipientsList()
            control.Model.StringItemList = recipients
            control.setText(recipients[0])

    def setSenders(self, senders):
        if len(senders) > 0:
            control = self._getSendersList()
            control.Model.StringItemList = senders
            control.selectItemPos(0, True)

    def setSaveSubject(self, state):
        self._getSaveSubject().Model.State = state

    def setSubject(self, subject):
        self._getSubject().Text = subject

    def setAttachmentAsPdf(self, state):
        self._getAttachmentAsPdf().Model.State = state

    def setSaveAttachments(self, state):
        self._getSaveAttachments().Model.State = state

    def setAttachments(self, attachments):
        if len(attachments):
            self._getAttachments().Model.StringItemList = attachments

    def enableButtonViewHtml(self):
        self._getButtonViewHtml().Model.Enabled = True

    def enableButtonViewPdf(self, enabled):
        self._getButtonViewPdf().Model.Enabled = enabled

    def addSender(self, sender):
        control = self._getSendersList()
        count = control.getItemCount()
        if sender not in control.getItems():
            control.addItem(sender, count)
            control.selectItemPos(count, True)

    def enableRemoveSender(self, enabled):
        self._getButtonRemoveSender().Model.Enabled = enabled

    def removeSender(self, position):
        control = self._getSendersList()
        control.removeItems(position, 1)
        if control.getItemCount() != 0:
            control.selectItemPos(0, True)

    def enableAddRecipient(self, enabled):
        self._getButtonAddRecipient().Model.Enabled = enabled

    def enableRemoveRecipient(self, enabled):
        self._getButtonRemoveRecipient().Model.Enabled = enabled

    def addRecipient(self):
        control = self._getRecipientsList()
        email = control.getText()
        self._addRecipient(control, email)

    def addToRecipient(self, email):
        control = self._getRecipientsList()
        self._addRecipient(control, email)

    def removeRecipient(self):
        self._getButtonRemoveRecipient().Model.Enabled = False
        control = self._getRecipientsList()
        email = control.getText()
        recipients = control.getItems()
        if email in recipients:
            control.setText('')
            position = recipients.index(email)
            control.removeItems(position, 1)

    def addAttachments(self, attachments):
        control = self._getAttachments()
        count = control.getItemCount()
        control.addItems(attachments, count)

    def enableRemoveAttachments(self, enabled):
        self._getButtonRemoveAttachments().Model.Enabled = enabled

    def removeAttachments(self):
        self._getButtonRemoveAttachments().Model.Enabled = False
        control = self._getAttachments()
        for position in reversed(control.getSelectedItemsPos()):
            control.removeItems(position, 1)

    def isDisposed(self):
        return self._window is None

    def dispose(self):
        self._window.dispose()
        self._window = None

# MailerView getter methods
    def getSender(self):
        control = self._getSendersList()
        sender = control.getSelectedItem()
        return sender

    def getRecipients(self):
        control = self._getRecipientsList()
        recipients = control.getItems()
        return recipients

    def getSelectedSender(self):
        control = self._getSendersList()
        sender = control.getSelectedItem()
        position = control.getSelectedItemPos()
        return sender, position

    def getAttachmentAsPdf(self):
        state = self._getAttachmentAsPdf().Model.State
        return bool(state)

    def getAttachments(self):
        attachments = self._getAttachments().getItems()
        return attachments

    def getSelectedAttachment(self):
        return self._getAttachments().getSelectedItem()

    def isSenderValid(self):
        return self._getSendersList().getSelectedItemPos() != -1

    def isRecipientsValid(self):
        return self._getRecipientsList().getItemCount() > 0

    def isSubjectValid(self):
        return self._getSubject().getText() != ''

    def getSubject(self):
        subject = self._getSubject().Text
        return subject

    def getSaveSubject(self):
        state = self._getSaveSubject().Model.State
        return bool(state)

    def getSaveAttachments(self):
        state = self._getSaveAttachments().Model.State
        return bool(state)

# MailerView private setter methods
    def _addRecipient(self, control, email):
        control.setText('')
        count = control.getItemCount()
        control.addItem(email, count)
        self._getButtonRemoveRecipient().Model.Enabled = False

# MailerView StringRessoure methods
    def getPropertyResource(self, name):
        return 'Mailer.Document.Property.%s' % name

    def getFilePickerTitleResource(self):
        return 'Mailer.FilePicker.Title'

# MailerView private control methods
    def _getSendersList(self):
        return self._window.getControl('ListBox1')

    def _getAttachments(self):
        return self._window.getControl('ListBox2')

    def _getRecipientsList(self):
        return self._window.getControl('ComboBox1')

    def _getSubject(self):
        return self._window.getControl('TextField1')

    def _getSaveSubject(self):
        return self._window.getControl('CheckBox1')

    def _getSaveAttachments(self):
        return self._window.getControl('CheckBox2')

    def _getAttachmentAsPdf(self):
        return self._window.getControl('CheckBox3')

    def _getButtonRemoveSender(self):
        return self._window.getControl('CommandButton2')

    def _getButtonAddRecipient(self):
        return self._window.getControl('CommandButton3')

    def _getButtonRemoveRecipient(self):
        return self._window.getControl('CommandButton4')

    def _getButtonViewHtml(self):
        return self._window.getControl('CommandButton5')

    def _getButtonRemoveAttachments(self):
        return self._window.getControl('CommandButton7')

    def _getButtonViewPdf(self):
        return self._window.getControl('CommandButton8')
