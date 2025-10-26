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

from ....dialog import MailView


class MergerView(MailView):
    def __init__(self, ctx, handler, parent, step):
        super().__init__(ctx, handler, parent, step)

# MergerView getter methods
    def getEmail(self):
        return self.getSender(), self.getRecipients(), self._getRecipientsFilter()

    def getSelectedRecipients(self):
        selection = []
        items = self._getRecipients().getSelectedItemsPos()
        for i in items:
            selection.append(i + 1)
        return tuple(selection)

    def getMergeAttachments(self):
        state = self._getMergeAttachments().Model.State
        return bool(state)

# MergerView setter methods
    def setRecipients(self, recipients, message):
        self._getMergerMessage().Text = message
        control = self._getRecipients()
        if control.Model.ItemCount:
            control.Model.StringItemList = ()
        for filter, email in recipients.items():
            index = control.Model.ItemCount
            control.Model.insertItemText(index, email)
            control.Model.setItemData(index, filter)
        if recipients:
            control.selectItemPos(0, True)

# MergerView private getter methods
    def _getRecipientsFilter(self):
        filters = []
        control = self._getRecipients()
        for index in range(control.Model.ItemCount):
            filters.append(control.Model.getItemData(index))
        return tuple(filters)

# MergerView private control methods
    def _getRecipients(self):
        return self._window.getControl('ListBox2')

    def _getMergeAttachments(self):
        return self._window.getControl('CheckBox3')
