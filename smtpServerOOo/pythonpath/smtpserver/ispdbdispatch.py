#!
# -*- coding: utf_8 -*-

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
from .wizardmodel import WizardModel
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
        else:
            self._parent = None
            print("IspdbDispatch.__init__() not parent set!!!!!!")
        logMessage(self.ctx, INFO, "Loading ... Done", 'IspdbDispatch', '__init__()')

    # XDispatch
    def dispatch(self, url, arguments):
        for arg in arguments:
            if arg.Name == 'Email':
                self._model.Email = arg.Value
        print("IspdbDispatch.dispatch()")
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
                msg +=  " Retrieving Authorization Code ..."
            else:
                msg +=  " ERROR: Wizard as been aborted"
            wizard.DialogWindow.dispose()
            print(msg)
            logMessage(self.ctx, INFO, msg, 'IspdbDispatch', '_showWizard()')
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)
