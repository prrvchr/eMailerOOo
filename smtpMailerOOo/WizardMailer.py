#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK
from com.sun.star.ui.dialogs.ExecutableDialogResults import CANCEL

from unolib import createService
from unolib import getStringResource

from smtpmailer import logMessage
from smtpmailer import getMessage
from smtpmailer import g_identifier
from smtpmailer import g_extension
from smtpmailer import g_wizard_paths
from smtpmailer import Wizard
from smtpmailer import WizardController

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = '%s.WizardMailer' % g_identifier


class WizardMailer(unohelper.Base,
                   XServiceInfo):
    def __init__(self, ctx):
        self.ctx = ctx
        msg = "Loading ... Done"
        #component = createService(self.ctx, 'com.sun.star.frame.Desktop').CurrentComponent
        #document = component.CurrentController.Frame
        #dispatcher = createService(self.ctx, 'com.sun.star.frame.DispatchHelper')
        #dispatcher.executeDispatch(document, '.uno:NewWindow', '', 0, ())
        #package = createService(self.ctx, 'com.sun.star.deployment.ui.PackageManagerDialog')

        #wizard = createService(self.ctx, 'com.sun.star.ui.dialogs.Wizard')
        wizard = Wizard(self.ctx)
        print("WizardMailer.__init__() 1")
        controller = WizardController(self.ctx, wizard)
        #arguments = ((uno.Any('[]short', g_wizard_paths), controller), )
        print("WizardMailer.__init__() 2")
        arguments = (g_wizard_paths, controller)
        #uno.invoke(wizard, 'initialize', arguments)
        wizard.initialize(arguments)
        print("WizardMailer.__init__() 3")
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)
        logMessage(self.ctx, INFO, msg, 'WizardMailer', '__init__()')
        print(msg)
        #mri = createService(self.ctx, 'mytools.Mri')
        #mri.inspect(package)
        wizard.execute()

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)


g_ImplementationHelper.addImplementation(WizardMailer,                              # UNO object class
                                         g_ImplementationName,                      # Implementation name
                                        (g_ImplementationName,))                    # List of implemented services
