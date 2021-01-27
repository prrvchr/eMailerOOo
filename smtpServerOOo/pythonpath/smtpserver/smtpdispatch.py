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

from com.sun.star.frame import XNotifyingDispatch

from com.sun.star.frame.DispatchResultState import SUCCESS
from com.sun.star.frame.DispatchResultState import FAILURE

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import getFileSequence
from unolib import getStringResource
from unolib import getResourceLocation
from unolib import getDialog
from unolib import createService

from .wizard import Wizard
from .wizardcontroller import WizardController

from .pagemodel import PageModel

from .spooler import SpoolerManager

from .configuration import g_extension
from .configuration import g_identifier
from .configuration import g_wizard_page
from .configuration import g_wizard_paths

from .logger import logMessage
from .logger import getMessage

import traceback


class SmtpDispatch(unohelper.Base,
                   XNotifyingDispatch):
    def __init__(self, ctx, url, parent):
        self._ctx = ctx
        self._parent = parent
        self._listeners = []
        print("SmtpDispatch.__init__()")

    # XNotifyingDispatch
    def dispatchWithNotification(self, url, arguments, listener):
        state = FAILURE
        result = None
        print("SmtpDispatch.dispatchWithNotification() 1")
        if url.Path == '//server':
            state, result = self._showSmtpServer()
        elif url.Path == '//spooler':
            state = SUCCESS
            self._showSmtpSpooler()
        struct = 'com.sun.star.frame.DispatchResultEvent'
        notification = uno.createUnoStruct(struct, self, state, result)
        print("SmtpDispatch.dispatchWithNotification() 2")
        listener.dispatchFinished(notification)

    def dispatch(self, url, arguments):
        print("SmtpDispatch.dispatch() 1")
        if url.Path == '//server':
            self._showSmtpServer()
        elif url.Path == '//spooler':
            self._showSmtpSpooler()
        print("SmtpDispatch.dispatch() 2")

    def addStatusListener(self, listener, url):
        print("SmtpDispatch.addStatusListener()")

    def removeStatusListener(self, listener, url):
        print("SmtpDispatch.removeStatusListener()")

    def _showSmtpServer(self):
        try:
            print("_showSmtpServer()")
            state = FAILURE
            email = None
            msg = "Wizard Loading ..."
            model = PageModel(self._ctx)
            wizard = Wizard(self._ctx, g_wizard_page, True, self._parent)
            controller = WizardController(self._ctx, wizard, model)
            arguments = (g_wizard_paths, controller)
            wizard.initialize(arguments)
            msg += " Done ..."
            if wizard.execute() == OK:
                state = SUCCESS
                email = model.Email
                msg +=  " Retrieving SMTP configuration OK..."
            wizard.DialogWindow.dispose()
            wizard.DialogWindow = None
            print(msg)
            logMessage(self._ctx, INFO, msg, 'SmtpDispatch', '_showSmtpServer()')
            return state, email
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def _showSmtpSpooler(self):
        try:
            print("SmtpDispatch._showSmtpSpooler() 1")
            manager = SpoolerManager(self._ctx)
            manager.viewSpooler(self._parent)
            print("SmtpDispatch._showSmtpSpooler() 2")
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)
