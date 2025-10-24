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

import unohelper

from com.sun.star.awt import XCallback

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from .ispdbview import IspdbView

from .ispdbhandler import WindowHandler

from .send import DialogHandler
from .send import SendView

from ....unotool import getArgumentSet
from ....unotool import getStringResource

from ....logger import LoggerListener

from ....configuration import g_identifier

import traceback


class IspdbManager(unohelper.Base,
                   XCallback):
    def __init__(self, ctx, wizard, model, pageid, parent):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._connected = False
        self._dialog = None
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'IspdbPage5')
        self._view = IspdbView(ctx, WindowHandler(self), parent)
        self._model.addLogListener(LoggerListener(self))
        self.updateLogger()

# XCallback
    def notify(self, data):
        self._notify(**getArgumentSet(data))

# XWizardPage
    @property
    def PageId(self):
        return self._pageid
    @property
    def Window(self):
        return self._view.getWindow()

    def activatePage(self):
        self._connected = False
        self._view.setPageStep(1)
        self._model.connectServers(self)

    def commitPage(self, reason):
        if reason == FINISH:
            self._model.saveConfiguration()
        return True

    def canAdvance(self):
        return self._connected

# IspdbManager setter methods
    def updateLogger(self):
        self._view.updateLogger(*self._model.getLogContent())

    def sendMail(self):
        parent = self._wizard.DialogWindow.getPeer()
        self._dialog = SendView(self._ctx, DialogHandler(self), parent, *self._model.getSendMailData())
        if self._dialog.execute() == OK:
            self._view.setPageStep(1)
            self._model.sendMessage(self, *self._getSendMailData())
        self._dialog.dispose()
        self._dialog = None

    def updateDialog(self):
        if self._dialog is not None:
            enabled = self._model.validSend(*self._getSendMailData())
            self._dialog.enableButtonSend(enabled)

# IspdbManager private methods
    def _getSendMailData(self):
        recipient = self._dialog.getRecipient()
        subject = self._dialog.getSubject()
        msg = self._dialog.getMessage()
        return recipient, subject, msg

    def _notify(self, call, **kwargs):
        if call == 'reset':
            self._resetProgress(**kwargs)
        elif call == 'progress':
            self._updateProgress(**kwargs)
        elif call == 'label':
            self._setLabel(**kwargs)
        elif call == 'step':
            self._setStep(**kwargs)

    def _resetProgress(self, value):
        if not self._model.isDisposed():
            self._view.resetProgress(value)

    def _updateProgress(self, value):
        if not self._model.isDisposed():
            self._view.updateProgress(value)

    def _setLabel(self, service, host, port):
        if not self._model.isDisposed():
            label = self._model.getPageLabel(self._resolver, self._pageid, service, host, port)
            self._view.setPageLabel(label)

    def _setStep(self, value):
        if not self._model.isDisposed():
            self._connected = value > 3
            self._view.setPageStep(value)
            self._wizard.updateTravelUI()

