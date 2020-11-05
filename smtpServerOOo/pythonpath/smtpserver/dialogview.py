#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from unolib import getDialog
from unolib import getStringResource

from .configuration import g_identifier
from .configuration import g_extension

from .logger import logMessage

import traceback


class DialogView(unohelper.Base):
    def __init__(self, ctx, xdl, handler, parent):
        self.ctx = ctx
        self._dialog = getDialog(self.ctx, g_extension, xdl, handler, parent)
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)
        print("DialogView.__init__()")

# DialogView setter methods
    def setTitle(self, context):
        server, port = context.getValueByName('ServerName'), context.getValueByName('Port')
        title = self._stringResource.resolveString(self._getTitleMessage())
        self._dialog.setTitle(title % (server, port))

    def updateProgress(self, value, offset=0, msg=None):
        self._getProgressBar().Value = value
        text = self._stringResource.resolveString(self._getProgressMessage(value + offset))
        if msg is not None:
            text = text % msg
        self._getProgressLabel().Text = text

    def callBack(self, state):
        if state:
            self._getButtonOk().Model.Enabled = True
        self._getButtonRetry().Model.Enabled = True

    def enableButtonOk(self, enabled):
        self._getButtonOk().Model.Enabled = enabled

    def enableButtonRetry(self, enabled):
        self._getButtonRetry().Model.Enabled = enabled

# DialogView getter methods
    def execute(self):
        return self._dialog.execute()

# DialogView private message methods
    def _getTitleMessage(self):
        return 'SmtpDialog.Title'

    def _getProgressMessage(self, value):
        return 'SmtpDialog.Label1.Label.%s' % value

# DialogView private control methods
    def _getProgressBar(self):
        return self._dialog.getControl('ProgressBar1')

    def _getProgressLabel(self):
        return self._dialog.getControl('Label1')

    def _getButtonRetry(self):
        return self._dialog.getControl('CommandButton1')

    def _getButtonOk(self):
        return self._dialog.getControl('CommandButton3')
