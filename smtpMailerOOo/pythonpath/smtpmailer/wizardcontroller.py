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

from com.sun.star.ui.dialogs import XWizardController
from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.awt import XCallback
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK
from com.sun.star.ui.dialogs.ExecutableDialogResults import CANCEL
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import PropertySet
from unolib import createService
from unolib import getConfiguration
from unolib import getDialogUrl
from unolib import getStringResource
from unolib import getContainerWindow

from .configuration import g_identifier
from .configuration import g_extension
from .configuration import g_column_index
from .configuration import g_column_filters
from .configuration import g_wizard_paths

from .wizardhandler import WizardHandler
from .wizardpage import WizardPage
from .logger import logMessage

import traceback

MOTOBIT = 1024 * 1024

class WizardController(unohelper.Base,
                       XWizardController):
    def __init__(self, ctx, wizard):
        self.ctx = ctx
        self._wizard = wizard
        self._provider = createService(self.ctx, 'com.sun.star.awt.ContainerWindowProvider')
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)
        self._configuration = getConfiguration(self.ctx, g_identifier, True)
        self._handler = WizardHandler(self.ctx, self._wizard)
        self._maxsize = self._configuration.getByName("MaxSizeMo") * MOTOBIT

    # XWizardController
    def createPage(self, parent, id):
        msg = "PageId: %s ..." % id
        url = getDialogUrl(g_extension, 'PageWizard%s' % id)
        window = self._provider.createContainerWindow(url, 'NotUsed', parent, self._handler)
        page = WizardPage(self.ctx, id, window, self._handler)
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardController', 'createPage()')
        return page

    def getPageTitle(self, id):
        return self._stringResource.resolveString('PageWizard%s.Step' % id)

    def canAdvance(self):
        return self._wizard.getCurrentPage().canAdvance()

    def onActivatePage(self, id):
        msg = "PageId: %s..." % id
        title = self._stringResource.resolveString('PageWizard%s.Title' % id)
        self._wizard.setTitle(title)
        backward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.PREVIOUS')
        forward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.NEXT')
        finish = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.FINISH')
        self._wizard.updateTravelUI()
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardController', 'onActivatePage()')

    def onDeactivatePage(self, id):
        if id == 1:
            pass
        elif id == 2:
            pass
        elif id == 3:
            pass

    def confirmFinish(self):
        return True
