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

from com.sun.star.frame import XDispatch
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

from .configuration import g_extension
from .configuration import g_identifier
from .configuration import g_wizard_page
from .configuration import g_wizard_paths

from .logger import logMessage
from .logger import getMessage

import traceback


class IspdbDispatch(unohelper.Base,
                    XDispatch):
    def __init__(self, ctx, model, frame):
        self.ctx = ctx
        self._model = model
        self._listeners = []
        if frame is not None:
            self._parent = frame.getContainerWindow()
            print("IspdbDispatch.__init__()")
        else:
            self._parent = None
            print("IspdbDispatch.__init__() not parent set!!!!!!")
        logMessage(self.ctx, INFO, "Loading ... Done", 'IspdbDispatch', '__init__()')

    # XDispatch
    def dispatch(self, url, arguments):
        for arg in arguments:
            if arg.Name == 'Email':
                self._model.Email = arg.Value
                print("IspdbDispatch.dispatch() 1 %s" % self._model.Email)
        print("IspdbDispatch.dispatch() 2")
        self._showWizard()
    def addStatusListener(self, listener, url):
        print("IspDBServer.addStatusListener()")
    def removeStatusListener(self, listener, url):
        print("IspDBServer.removeStatusListener()")

    def _showWizard(self):
        try:
            print("_showWizard()")
            msg = "Wizard Loading ..."
            wizard = Wizard(self.ctx, g_wizard_page, True, self._parent)
            controller = WizardController(self.ctx, wizard, self._model)
            arguments = (g_wizard_paths, controller)
            wizard.initialize(arguments)
            msg += " Done ..."
            if wizard.execute() == OK:
                msg +=  " Retrieving SMTP configuration OK..."
            else:
                msg +=  " ERROR: Wizard as been aborted"
            wizard.DialogWindow.dispose()
            wizard.DialogWindow = None
            print(msg)
            logMessage(self.ctx, INFO, msg, 'IspdbDispatch', '_showWizard()')
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)
