#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.ui.dialogs import XWizardController
from com.sun.star.awt import XCallback
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK
from com.sun.star.ui.dialogs.ExecutableDialogResults import CANCEL

from unolib import PropertySet
from unolib import createService
from unolib import generateUuid
from unolib import getCurrentLocale
from unolib import getProperty
from unolib import getStringResource
from unolib import getContainerWindow
from unolib import getDialogUrl

from smtpmailer import logMessage
from smtpmailer import getMessage
from smtpmailer import g_identifier
from smtpmailer import g_extension
from smtpmailer import g_wizard_paths
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
        wizard = createService(self.ctx, 'com.sun.star.ui.dialogs.Wizard')
        controller = WizardController(self.ctx, wizard)
        arguments = ((uno.Any('[][]short', g_wizard_paths), controller), )
        uno.invoke(wizard, 'initialize', arguments)
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)
        logMessage(self.ctx, INFO, msg, 'WizardMailer', '__init__()')
        print(msg)
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
