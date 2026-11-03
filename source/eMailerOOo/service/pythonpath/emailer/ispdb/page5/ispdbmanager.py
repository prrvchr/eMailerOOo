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

from com.sun.star.mail.MailServiceType import IMAP
from com.sun.star.mail.MailServiceType import SMTP

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from .ispdbview import IspdbView

from .ispdbhandler import WindowHandler

from .send import DialogHandler
from .send import SendView

from ...unotool import getStringResource

from ...logger import LoggerListener

from ...configuration import g_identifier


class IspdbManager(unohelper.Base):
    def __init__(self, ctx, wizard, model, pageid, parent):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._connected = False
        self._dialog = None
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'IspdbPage5')
        self._view = IspdbView(ctx, WindowHandler(self), parent)
        self._model.addLogListener(LoggerListener(self.updateLogger))
        self.updateLogger()

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
        self._setLabel(SMTP.value)
        indicator = self._view.getIndicator(finish=False)
        self._model.connectServer(self.onConnectServer, indicator, SMTP)

    def commitPage(self, reason):
        if reason == FINISH:
            self._model.saveConfiguration()
        return True

    def canAdvance(self):
        return self._connected

# IspdbManager setter methods
    def onConnectServer(self, _success, connect, imap, _error):
        if imap:
            self._setLabel(IMAP.value)
            indicator = self._view.getIndicator(start=False)
            self._model.connectServer(self.onConnectServer, indicator, IMAP)
        else:
            self._setStep(4 if connect else 2)

    def updateLogger(self):
        self._view.updateLogger(*self._model.getLogContent())

    def sendMail(self):
        parent = self._wizard.DialogWindow.getPeer()
        self._dialog = SendView(self._ctx, DialogHandler(self), parent, *self._model.getSendMailData())
        if self._dialog.execute() == OK:
            self._view.setPageStep(1)
            indicator = self._view.getIndicator(finish=False)
            imap = self._model.useIMAP()
            self._model.sendMail(self.onSendMail, indicator, imap, *self._getSendMailData())
        else:
            self._dialog.dispose()
            self._dialog = None

    def onSendMail(self, success, mail, _error):
        if mail:
            indicator = self._view.getIndicator(start=False)
            thread = mail.ForeignId if mail.ForeignId else mail.MessageId
            self._model.sendMail(self.onSendMail, indicator, False, *self._getSendMailData(thread))
        else:
            if self._dialog:
                self._dialog.dispose()
            self._dialog = None
            self._setStep(5 if success else 3)

    def updateDialog(self):
        if self._dialog is not None:
            enabled = self._model.validSend(*self._getSendMailData())
            self._dialog.enableButtonSend(enabled)

# IspdbManager private methods
    def _getSendMailData(self, thread=None):
        recipient = self._dialog.getRecipient()
        subject = self._dialog.getSubject()
        msg = self._dialog.getMessage()
        return recipient, subject, msg, thread

    def _setLabel(self, service):
        if not self._model.isDisposed():
            self._view.setPageLabel(self._model.getMailServerLabel(self._resolver, self._pageid, service))

    def _setStep(self, value):
        if not self._model.isDisposed():
            self._connected = value > 3
            self._view.setPageStep(value)
            self._wizard.updateTravelUI()
