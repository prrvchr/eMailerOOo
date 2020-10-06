#!
# -*- coding: utf_8 -*-

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

from .configuration import g_identifier
from .configuration import g_extension

from .wizardhandler import WizardHandler
from .wizardpage import WizardPage
from .logger import logMessage

import traceback


class WizardController(unohelper.Base,
                       XWizardController):
    def __init__(self, ctx, wizard):
        self.ctx = ctx
        self._wizard = wizard
        self._provider = createService(self.ctx, 'com.sun.star.awt.ContainerWindowProvider')
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)
        self._configuration = getConfiguration(self.ctx, g_identifier, True)
        self._handler = WizardHandler(self.ctx, self._wizard)

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
