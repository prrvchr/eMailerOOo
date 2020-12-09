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

from com.sun.star.lang import XServiceInfo

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK
from com.sun.star.ui.dialogs.ExecutableDialogResults import CANCEL

from unolib import createService
from unolib import getStringResource
from unolib import getParentWindow

from smtpmailer import logMessage
from smtpmailer import getMessage
from smtpmailer import g_identifier
from smtpmailer import g_extension
from smtpmailer import g_wizard_paths
from smtpmailer import g_wizard_page
from smtpmailer import Wizard
from smtpmailer import WizardController
from smtpmailer import PageModel

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = '%s.WizardMailer' % g_identifier


class WizardMailer(unohelper.Base,
                   XServiceInfo):
    def __init__(self, ctx):
        try:
            print("WizardMailer.__init__() 1")
            self.ctx = ctx
            msg = "Loading ... Done"
            #component = createService(self.ctx, 'com.sun.star.frame.Desktop').CurrentComponent
            #document = component.CurrentController.Frame
            #dispatcher = createService(self.ctx, 'com.sun.star.frame.DispatchHelper')
            #dispatcher.executeDispatch(document, '.uno:NewWindow', '', 0, ())
            #package = createService(self.ctx, 'com.sun.star.deployment.ui.PackageManagerDialog')
    
            #wizard = createService(self.ctx, 'com.sun.star.ui.dialogs.Wizard')
            parent = getParentWindow(self.ctx)
            wizard = Wizard(self.ctx, g_wizard_page, True, parent)
            print("WizardMailer.__init__() 2")
            model = PageModel(self.ctx)
            controller = WizardController(self.ctx, wizard, model)
            #arguments = ((uno.Any('[]short', g_wizard_paths), controller), )
            print("WizardMailer.__init__() 3")
            arguments = (g_wizard_paths, controller)
            #uno.invoke(wizard, 'initialize', arguments)
            wizard.initialize(arguments)
            print("WizardMailer.__init__() 4")
            logMessage(self.ctx, INFO, msg, 'WizardMailer', '__init__()')
            print(msg)
            #mri = createService(self.ctx, 'mytools.Mri')
            #mri.inspect(package)
            if wizard.execute() == OK:
                print(" Retrieving SMTP configuration OK...")
            wizard.DialogWindow.dispose()
            wizard.DialogWindow = None
        except Exception as e:
            print("WizardMailer.__init__() ERROR: %s - %s" % (e, traceback.print_exc()))

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
