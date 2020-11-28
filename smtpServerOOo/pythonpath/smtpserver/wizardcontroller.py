#!
# -*- coding: utf_8 -*-

'''
    Copyright (c) 2020 https://prrvchr.github.io

    Permission is hereby granted, free of charge, to any person obtaining
    a copy of this software and associated documentation files (the "Software"),
    to deal in the Software without restriction, including without limitation
    the rights to use, copy, modify, merge, publish, distribute, sublicense,
    and/or sell copies of the Software, and to permit persons to whom the Software
    is furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
    EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
    OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
    CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
    TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
    OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'''

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizardController

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import createService
from unolib import getDialogUrl

from .pagemanager import PageManager
from .pagehandler import PageHandler
from .wizardpage import WizardPage

from .configuration import g_extension

from .logger import logMessage

import traceback


class WizardController(unohelper.Base,
                       XWizardController):
    def __init__(self, ctx, wizard, model=None):
        self.ctx = ctx
        #self._provider = createService(self.ctx, 'com.sun.star.awt.ContainerWindowProvider')
        manager = PageManager(self.ctx, wizard, model)
        self._handler = PageHandler(manager)

    # XWizardController
    def createPage(self, parent, pageid):
        msg = "PageId: %s ..." % pageid
        print("WizardController.createPage() 1")
        url = getDialogUrl(g_extension, 'PageWizard%s' % pageid)
        print("WizardController.createPage() 2")
        provider = createService(self.ctx, 'com.sun.star.awt.ContainerWindowProvider')
        print("WizardController.createPage() 3")
        try:
            window = provider.createContainerWindow(url, 'NotUsed', parent, self._handler)
            print("WizardController.createPage() 4")
        except Exception as e:
            msg = "WizardController.createPage() Error: %s - %s" % (e, traceback.print_exc())
            print(msg)
        # TODO: Fixed: When initializing XWizardPage, the handler must be disabled...
        self._handler.disable()
        print("WizardController.createPage() 5")
        page = WizardPage(self.ctx, pageid, window, self._handler.Manager)
        print("WizardController.createPage() 6")
        self._handler.enable()
        print("WizardController.createPage() 7")
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardController', 'createPage()')
        return page

    def getPageTitle(self, pageid):
        return self._handler.Manager.Wizard._manager.getPageStep(pageid)

    def canAdvance(self):
        return True

    def onActivatePage(self, pageid):
        msg = "PageId: %s..." % pageid
        self._handler.Manager.setPageTitle(pageid)
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
