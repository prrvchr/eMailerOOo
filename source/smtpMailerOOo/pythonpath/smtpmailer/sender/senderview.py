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

import uno
import unohelper

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from smtpmailer import createService
from smtpmailer import getDialog
from smtpmailer import g_extension

from .senderhandler import DialogHandler

import traceback


class SenderView(unohelper.Base):
    def __init__(self, ctx, manager, parent):
        handler = DialogHandler(manager)
        self._dialog = getDialog(ctx, g_extension, 'SenderDialog', handler, parent)

# SenderView setter methods
    def setTitle(self, title):
        self._dialog.setTitle(title)

    def enableButtonSend(self, enabled):
        self._getButtonSend().Model.Enabled = enabled

    def endDialog(self):
        self._dialog.endDialog(OK)

    def dispose(self):
        self._dialog.dispose()

# SenderView getter methods
    def getParent(self):
        return self._dialog.getPeer()

    def execute(self):
        return self._dialog.execute()

# SenderView private control methods
    def _getButtonSend(self):
        return self._dialog.getControl('CommandButton2')
