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

from ...mail import MailView


class MergerView(MailView):
# MergerView getter methods
    def getWindow(self):
        return self._window

    def isRecipientsValid(self):
        return self._getMergerRecipients().getItemCount() > 0

    def getRecipients(self):
        recipients = []
        identifiers = []
        control = self._getMergerRecipients()
        for index in range(control.Model.ItemCount):
            recipients.append(control.Model.getItemText(index))
            identifiers.append(control.Model.getItemData(index))
        return tuple(recipients), tuple(identifiers)

    def getCurrentIdentifier(self):
        identifier = None
        control = self._getMergerRecipients()
        index = control.getSelectedItemPos()
        if index != -1:
            identifier = control.Model.getItemData(index)
        return identifier

# MergerView setter methods
    def setMergerRecipient(self, recipients, message):
        self._getMergerMessage().Text = message
        control = self._getMergerRecipients()
        control.Model.removeAllItems()
        for recipient in recipients:
            index = control.Model.ItemCount
            control.Model.insertItemText(index, recipient.Recipient)
            control.Model.setItemData(index, recipient.Identifier)
        if len(recipients) > 0:
            control.selectItemPos(0, True)
