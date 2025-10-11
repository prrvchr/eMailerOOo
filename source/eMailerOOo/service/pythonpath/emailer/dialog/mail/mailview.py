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

import unohelper

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from ...unotool import createService
from ...unotool import getContainerWindow

from ...configuration import g_identifier

class MailView(unohelper.Base):
    def __init__(self, ctx, handler, parent, step):
        self._ctx = ctx
        self._window = getContainerWindow(ctx, parent, handler, g_identifier, 'MailWindow')
        self._window.Model.Step = step
        self._window.setVisible(True)

# MailView getter methods
    def getEmail(self):
        raise NotImplementedError('Need to be implemented!')

    def getSelectedRecipients(self):
        raise NotImplementedError('Need to be implemented!')

    def getWindow(self):
        return self._window

    def getSender(self):
        return self._getSenders().getSelectedItem()

    def getRecipients(self):
        return self._getRecipients().getItems()

    def getSelectedSender(self):
        control = self._getSenders()
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

    def hasSelectedPdfAttachment(self):
        control = self._getAttachments()
        position = control.getSelectedItemPos()
        print("MailView.hasSelectedPdfAttachment() position: %s" % position)
        return position != -1 and control.getSelectedItem().endswith('pdf')

    def getSubject(self):
        return self._getSubject().Text.strip()

    def getSaveSubject(self):
        state = self._getSaveSubject().Model.State
        return bool(state)

    def getSaveAttachments(self):
        state = self._getSaveAttachments().Model.State
        return bool(state)

    def getMergeAttachments(self):
        state = self._getMergeAttachments().Model.State
        return bool(state)

# MailView setter methods
    def setSubjectBackground(self, color):
        self._getSubject().Model.BackgroundColor = color

    def enableRemoveSender(self, enabled):
        self._getRemoveSender().Model.Enabled = enabled

    def setSenders(self, senders):
        if len(senders) > 0:
            control = self._getSenders()
            control.Model.StringItemList = senders
            control.selectItemPos(0, True)

    def addSender(self, sender):
        control = self._getSenders()
        count = control.getItemCount()
        if sender not in control.getItems():
            control.addItem(sender, count)
            control.selectItemPos(count, True)

    def removeSender(self, position):
        control = self._getSenders()
        control.removeItems(position, 1)
        if control.getItemCount() != 0:
            control.selectItemPos(0, True)

    def setSubject(self, subject):
        self._getSubject().Text = subject

    def setSaveSubject(self, state):
        self._getSaveSubject().Model.State = state

    def enableViewHtml(self, enabled):
        self._getViewHtml().Model.Enabled = enabled

    def setMessage(self, text):
        self._getMessage().Text = text

    def enableAddRecipient(self, enabled):
        self._getAddRecipient().Model.Enabled = enabled

    def enableRemoveRecipient(self, enabled):
        self._getRemoveRecipient().Model.Enabled = enabled

    def setRecipients(self, recipients):
        raise NotImplementedError('Need to be implemented if needed!')

    def enableViewPdf(self, enabled):
        self._getViewPdf().Model.Enabled = enabled

    def enableRemoveAttachments(self, enabled):
        self._getRemoveAttachments().Model.Enabled = enabled

    def enableMoveBefore(self, enabled):
        self._getMoveBefore().Model.Enabled = enabled

    def enableMoveAfter(self, enabled):
        self._getMoveAfter().Model.Enabled = enabled

    def setAttachmentAsPdf(self, state):
        self._getAttachmentAsPdf().Model.State = state

    def setSaveAttachments(self, state):
        self._getSaveAttachments().Model.State = state

    def setAttachments(self, attachments):
        if len(attachments):
            self._getAttachments().Model.StringItemList = attachments

    def addAttachments(self, attachments):
        control = self._getAttachments()
        count = control.getItemCount()
        control.addItems(attachments, count)

    def moveAttachments(self, offset):
        control = self._getAttachments()
        positions = control.getSelectedItemsPos()
        attachments = control.getSelectedItems()
        for position in reversed(positions):
            control.Model.removeItem(position)
        positions = tuple([position + offset for position in positions])
        for index, position in enumerate(positions):
            attachment = attachments[index]
            control.Model.insertItemText(position, attachment)
        control.selectItemsPos(positions, True)

    def removeAttachments(self):
        self._getRemoveAttachments().Model.Enabled = False
        self._getViewPdf().Model.Enabled = False
        control = self._getAttachments()
        for position in reversed(control.getSelectedItemsPos()):
            control.removeItems(position, 1)

# MailView private setter methods
    def _addRecipient(self, control, email):
        control.setText('')
        count = control.getItemCount()
        control.addItem(email, count)
        self._getRemoveRecipient().Model.Enabled = False

# MailView private control methods
    def _getSenders(self):
        return self._window.getControl('ListBox1')

    def _getRecipients(self):
        raise NotImplementedError('Need to be implemented if needed!')

    def _getAttachments(self):
        return self._window.getControl('ListBox3')

    def _getSubject(self):
        return self._window.getControl('TextField1')

    def _getSaveSubject(self):
        return self._window.getControl('CheckBox1')

    def _getAttachmentAsPdf(self):
        return self._window.getControl('CheckBox2')

    def _getSaveAttachments(self):
        return self._window.getControl('CheckBox4')

    def _getRemoveSender(self):
        return self._window.getControl('CommandButton2')

    def _getViewHtml(self):
        return self._window.getControl('CommandButton5')

    def _getRemoveAttachments(self):
        return self._window.getControl('CommandButton7')

    def _getMoveBefore(self):
        return self._window.getControl('CommandButton8')

    def _getMoveAfter(self):
        return self._window.getControl('CommandButton9')

    def _getViewPdf(self):
        return self._window.getControl('CommandButton10')

    def _getMessage(self):
        return self._window.getControl('Label8')

    def _getMergerMessage(self):
        return self._window.getControl('Label4')
