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

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from unolib import getDialog
from unolib import createService

from .senderhandler import DialogHandler

from smtpserver import g_extension

import traceback


class SenderView(unohelper.Base):
    def __init__(self, ctx):
        self._ctx = ctx
        self._dialog = None

# SenderView setter methods
    def setDialog(self, manager, parent):
        handler = DialogHandler(manager)
        self._dialog = getDialog(self._ctx, g_extension, 'SenderDialog', handler, parent)

    def setTitle(self, title):
        self._dialog.setTitle(title)

    def enableButtonSend(self, enabled):
        self._getButtonSend().Model.Enabled = enabled

    def endDialog(self):
        self._dialog.endDialog(OK)

    def dispose(self):
        self._dialog.dispose()
        self._dialog = None

# SenderView getter methods
    def getDocumentUrlAndPath(self, path, title, filter):
        url = None
        service = 'com.sun.star.ui.dialogs.FilePicker'
        filepicker = createService(self._ctx, service)
        filepicker.setTitle(title)
        filepicker.setDisplayDirectory(path)
        filepicker.appendFilter(*filter)
        filepicker.setCurrentFilter(filter[0])
        if filepicker.execute() == OK:
            url = filepicker.getSelectedFiles()[0]
            path = filepicker.getDisplayDirectory()
        filepicker.dispose()
        return url, path

    def getParent(self):
        return self._dialog.getPeer()

    def execute(self):
        return self._dialog.execute()

    def isDisposed(self):
        return self._dialog is None

# SenderView StringRessoure methods
    def getTitleRessource(self):
        return 'SenderDialog.Title'

    def getFilePickerTitleResource(self):
        return 'Sender.FilePicker.Title'

    def getFilePickerFilterResource(self):
        return 'Sender.FilePicker.Filter.Writer'

# SenderView private control methods
    def _getButtonSend(self):
        return self._dialog.getControl('CommandButton2')
