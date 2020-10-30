#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizardController

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService
from unolib import getDialogUrl

from .wizardpage import WizardPage
from .pagehandler import PageHandler

from .configuration import g_extension

from .logger import logMessage

import traceback


class WizardController(unohelper.Base,
                       XWizardController):
    def __init__(self, ctx, wizard, model=None):
        self.ctx = ctx
        self._provider = createService(self.ctx, 'com.sun.star.awt.ContainerWindowProvider')
        self._handler = PageHandler(self.ctx, wizard, model)

    # XWizardController
    def createPage(self, parent, pageid):
        msg = "PageId: %s ..." % pageid
        url = getDialogUrl(g_extension, 'PageWizard%s' % pageid)
        window = self._provider.createContainerWindow(url, 'NotUsed', parent, self._handler)
        # TODO: Fixed: When initializing XWizardPage, the handler must be disabled...
        self._handler.disable()
        page = WizardPage(self.ctx, pageid, window, self._handler.getManager())
        self._handler.enable()
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardController', 'createPage()')
        return page

    def getPageTitle(self, pageid):
        return self._handler.getManager().getPageStep(pageid)

    def canAdvance(self):
        return True

    def onActivatePage(self, pageid):
        msg = "PageId: %s..." % pageid
        self._handler.getManager().setPageTitle(pageid)
        backward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.PREVIOUS')
        forward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.NEXT')
        finish = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.FINISH')
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardController', 'onActivatePage()')

    def onDeactivatePage(self, pageid):
        if pageid == 1:
            pass
        elif pageid == 2:
            pass
        elif pageid == 3:
            pass

    def confirmFinish(self):
        return True
