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

import unohelper

from smtpmailer import getDialog
from smtpmailer import g_extension

from .sendhandler import DialogHandler


class SendView(unohelper.Base):
    def __init__(self, ctx, manager, parent, title, email, subject, msg):
        handler = DialogHandler(manager)
        self._dialog = getDialog(ctx, g_extension, 'SendDialog', handler, parent)
        self._dialog.setTitle(title)
        self._getRecipient().Text = email
        self._getSubject().Text = subject
        self._getMessage().Text = msg

# DialogView setter methods
    def enableButtonSend(self, enabled):
        self._getButtonSend().Model.Enabled = enabled

    def dispose(self):
        self._dialog.dispose()

# DialogView getter methods
    def execute(self):
        return self._dialog.execute()

    def getRecipient(self):
        return self._getRecipient().Text

    def getSubject(self):
        return self._getSubject().Text

    def getMessage(self):
        return self._getMessage().Text

# DialogView private control methods
    def _getRecipient(self):
        return self._dialog.getControl('TextField1')

    def _getSubject(self):
        return self._dialog.getControl('TextField2')

    def _getMessage(self):
        return self._dialog.getControl('TextField3')

    def _getButtonSend(self):
        return self._dialog.getControl('CommandButton2')
