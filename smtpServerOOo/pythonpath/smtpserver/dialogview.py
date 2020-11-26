#!
# -*- coding: utf_8 -*-

import unohelper

from unolib import getDialog

from .configuration import g_extension


class DialogView(unohelper.Base):
    def __init__(self, ctx, xdl, handler, parent):
        self._dialog = getDialog(ctx, g_extension, xdl, handler, parent)

# DialogView setter methods
    def setTitle(self, model):
        title = model.resolveString(self._getTitle())
        self._dialog.setTitle(title % model.Email)

    def enableButtonSend(self, model):
        enabled = all((model.isEmailValid(self.getRecipient()),
                       model.isStringValid(self.getObject()),
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

    def getObject(self):
        return self._dialog.getControl('TextField2').Text

    def getMessage(self):
        return self._dialog.getControl('TextField3').Text

# DialogView private message methods
    def _getTitle(self):
        return 'SendDialog.Title'

# DialogView private control methods
    def _getButtonSend(self):
        return self._dialog.getControl('CommandButton2')
