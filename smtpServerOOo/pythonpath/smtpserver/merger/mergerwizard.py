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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpserver import logMessage

from .mergermodel import MergerModel

from .page1 import MergerManager as WizardPage1
from .page2 import MergerManager as WizardPage2
from .page3 import MergerManager as WizardPage3

import traceback


class MergerWizard(unohelper.Base,
                   XWizardController):
    def __init__(self, ctx, wizard, datasource):
        self._ctx = ctx
        self._wizard = wizard
        self._model = MergerModel(ctx, datasource)

    def dispose(self):
        self._model.dispose()
        self._wizard.DialogWindow.dispose()

# XWizardController
    def createPage(self, parent, pageid):
        try:
            msg = "PageId: %s ..." % pageid
            if pageid == 1:
                page = WizardPage1(self._ctx, self._wizard, self._model, pageid, parent)
            elif pageid == 2:
                page = WizardPage2(self._ctx, self._wizard, self._model, pageid, parent)
            elif pageid == 3:
                page = WizardPage3(self._ctx, self._wizard, self._model, pageid, parent)
            msg += " Done"
            logMessage(self._ctx, INFO, msg, 'WizardController', 'createPage()')
            return page
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def getPageTitle(self, pageid):
        return self._model.getPageStep(pageid)

    def canAdvance(self):
        return True

    def onActivatePage(self, pageid):
        msg = "PageId: %s..." % pageid
        title = self._model.getPageTitle(pageid)
        self._wizard.setTitle(title)
        backward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.PREVIOUS')
        forward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.NEXT')
        finish = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.FINISH')
        msg += " Done"
        logMessage(self._ctx, INFO, msg, 'WizardController', 'onActivatePage()')

    def onDeactivatePage(self, pageid):
        pass

    def confirmFinish(self):
        return True
