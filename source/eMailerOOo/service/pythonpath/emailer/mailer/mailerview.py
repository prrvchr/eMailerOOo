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

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from ..mail import MailView

from ..unotool import getDialog

from ..configuration import g_identifier


class MailerView(MailView):
    def __init__(self, ctx, handler1, handler2, parent, step):
        self._dialog = getDialog(ctx, g_identifier, 'MailerDialog', handler1, parent)
        MailView.__init__(self, ctx, handler2, self._dialog.getPeer(), step)

# MailerView getter methods
    def getRecipientIndex(self):
        return -1

    def getRecipient(self):
        return self._getRecipients().getText().strip()

# MailerView setter methods
    def execute(self):
        return self._dialog.execute()

    def setTitle(self, title):
        self._dialog.setTitle(title)

    def enableButtonSend(self, enabled):
        self._getButtonSend().Model.Enabled = enabled

    def endDialog(self):
        self._dialog.endDialog(OK)

    def dispose(self):
        self._dialog.dispose()

    def addRecipient(self, recipient):
        control = self._getRecipients()
        count = control.getItemCount()
        control.addItem(recipient, count)
        control.setText(recipient)
        self._getRemoveRecipient().Model.Enabled = False

    def removeRecipient(self):
        self._getRemoveRecipient().Model.Enabled = False
        control = self._getRecipients()
        email = control.getText()
        recipients = control.getItems()
        if email in recipients:
            control.setText('')
            position = recipients.index(email)
            control.removeItems(position, 1)


# MailerView private control methods
    def _getRecipients(self):
        return self._window.getControl('ComboBox1')

    def _getAddRecipient(self):
        return self._window.getControl('CommandButton3')

    def _getRemoveRecipient(self):
        return self._window.getControl('CommandButton4')

    def _getButtonSend(self):
        return self._dialog.getControl('CommandButton2')

