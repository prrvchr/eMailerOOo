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

from ...mail.mailview import MailView

from ...unotool.layout import Layout


class MergerView(MailView):
    def __init__(self, ctx, handler, parent, step):
        super().__init__(ctx, handler, parent, step)
        self._layout = Layout(self._window, **self._getLayouts())

# MergerView getter methods
    def getLayout(self):
        return self._layout

    def getEmail(self):
        return self.getSender(), self.getRecipients(), self._getRecipientsFilter()

    def hasSelectedRecipient(self):
        return self._getRecipients().getSelectedItemPos() != - 1

    def getSelectedRecipient(self):
        recipient = None
        control = self._getRecipients()
        index = control.getSelectedItemPos()
        if index != -1:
            recipient = control.Model.getItemData(index)
        return recipient

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
        for fltr, email in recipients.items():
            index = control.Model.ItemCount
            control.Model.insertItemText(index, email)
            control.Model.setItemData(index, fltr)
        if recipients:
            control.selectItemPos(0, True)

# MergerView private getter methods
    def _getLayouts(self):
        return {'horizontals': self._getHorizontals(),
                'controls': self._getControls(),
                'rows': self._getRows(),
                'tops': self._getTops()}

    def _getHorizontals(self):
        for index in range(1, 3):
            yield self._getFixedLine(index)
        yield self._getSenders()
        yield self._getRecipients()
        yield self._getSubject()

    def _getControls(self):
        yield self._getAttachments()

    def _getRows(self):
        yield self._getMergeAttachments
        yield self._getSaveAttachments()

    def _getTops(self):
        yield self._getAddSender()
        yield self._getRemoveSender()
        yield self._getMergerMessage()
        yield self._getSaveSubject()
        for index in (5, 12):
            yield self._getButtons(index)

    def _getRecipientsFilter(self):
        filters = []
        control = self._getRecipients()
        for index in range(control.Model.ItemCount):
            filters.append(control.Model.getItemData(index))
        return tuple(filters)

# MergerView private control methods
    def _getRecipients(self):
        return self._window.getControl('ListBox2')

    def _getFixedLine(self, index):
        return self._window.getControl(f'FixedLine{index}')

    def _getAddSender(self):
        return self._window.getControl('CommandButton1')

    def _getMergeAttachments(self):
        return self._window.getControl('CheckBox3')

    def _getButtons(self, index):
        return self._window.getControl(f'CommandButton{index}')
