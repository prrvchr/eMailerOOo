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
    def __init__(self, ctx, xdl):
        self.ctx = ctx
        self._dialog = getDialog(self.ctx, g_extension, xdl)
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)
        print("DialogView.__init__()")

# DialogView setter methods
    def setTitle(self, context):
        server, port = context.getValueByName('ServerName'), context.getValueByName('Port')
        title = self._stringResource.resolveString(self._getTitleMessage())
        self._dialog.setTitle(title % (server, port))

    def updateProgress(self, value, offset=0):
        self._getProgressBar().Value = value
        text = self._stringResource.resolveString(self._getProgressMessage(value + offset))
        self._getProgressLabel().Text = text

    def callBack(self):
        self._getButtonOk().Model.Enabled = True

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

    def _getButtonOk(self):
        return self._dialog.getControl('CommandButton2')
