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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService
from unolib import getUrl

from .mailermodel import MailerModel
from .mailerview import MailerView

from ..logger import logMessage
from ..logger import getMessage
g_message = 'mailermanager'

import traceback


class MailerManager(unohelper.Base):
    def __init__(self, ctx, datasource, url, parent):
        self._ctx = ctx
        self._model = MailerModel(ctx, datasource, url)
        self._view = MailerView(ctx, self, parent)
        print("MailerManager.__init__()")

    @property
    def Model(self):
        return self._model

    def show(self):
        return self._view.execute()

    def addSender(self):
        print("MailerManager.addSender()")
        try:
            print("MailerManager.addSender() 1")
            url = getUrl(self._ctx, 'ispdb://')
            desktop = createService(self._ctx, 'com.sun.star.frame.Desktop')
            dispatcher = desktop.getCurrentFrame().queryDispatch(url, '', 0)
            #dispatcher = createService(self._ctx, 'com.sun.star.frame.DispatchHelper')
            #dispatcher.executeDispatch(desktop.getCurrentFrame(), 'ispdb://', '', 0, ())
            print("MailerManager.addSender() 2")
            if dispatcher is not None:
                #dispatcher.dispatchWithNotification(url, (), self)
                dispatcher.dispatch(url, ())
                print("MailerManager.addSender() 3")
            #mri = createService(self._ctx, 'mytools.Mri')
            #mri.inspect(desktop)
            #msg = "OptionsDialog._showWizard()"
            #logMessage(self._ctx, INFO, msg, 'OptionsDialog', '_showWizard()')
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)



    def removeSender(self):
        print("MailerManager.removeSender()")

    def changeRecipient(self):
        print("MailerManager.changeRecipient()")

    def addRecipient(self):
        print("MailerManager.addRecipient()")

    def removeRecipient(self):
        print("MailerManager.removeRecipient()")
