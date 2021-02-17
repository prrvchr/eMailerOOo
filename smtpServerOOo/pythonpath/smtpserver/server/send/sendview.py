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

import unohelper

from unolib import getDialog

from .sendhandler import DialogHandler

from smtpserver.configuration import g_extension


class SendView(unohelper.Base):
    def __init__(self, ctx, manager, parent):
        handler = DialogHandler(manager)
        self._dialog = getDialog(ctx, g_extension, 'SendDialog', handler, parent)
        self._setTitle(manager.Model)
        self._getRecipient().Text = manager.Model.Email
        self._getSubject().Text = manager.Model.resolveString('SendDialog.TextField2.Text')
        self._getMessage().Text = manager.Model.resolveString('SendDialog.TextField3.Text')
        self.enableButtonSend(manager.Model)

# DialogView setter methods
    def enableButtonSend(self, model):
        enabled = all((model.isEmailValid(self.getRecipient()),
                       model.isStringValid(self.getSubject()),
                       model.isStringValid(self.getMessage())))
        self._getButtonSend().Model.Enabled = enabled

    def dispose(self):
        self._dialog.dispose()
        self._dialog = None

# DialogView getter methods
    def execute(self):
        return self._dialog.execute()

    def getRecipient(self):
        return self._dialog.getControl('TextField1').Text

    def getSubject(self):
        return self._dialog.getControl('TextField2').Text

    def getMessage(self):
        return self._dialog.getControl('TextField3').Text

# DialogView private setter methods
    def _setTitle(self, model):
        title = model.resolveString('SendDialog.Title')
        self._dialog.setTitle(title % model.Email)

# DialogView private control methods
    def _getButtonSend(self):
        return self._dialog.getControl('CommandButton2')

    def _getRecipient(self):
        return self._dialog.getControl('TextField1')

    def _getSubject(self):
        return self._dialog.getControl('TextField2')

    def _getMessage(self):
        return self._dialog.getControl('TextField3')
