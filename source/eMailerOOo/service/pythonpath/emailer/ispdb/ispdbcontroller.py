#!
# -*- coding: utf-8 -*-

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

from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.mail.MailServiceType import IMAP

from com.sun.star.ui.dialogs import XWizardController

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from ..logger import logMessage

from .ispdbmodel import IspdbModel

from .page1 import IspdbManager as WizardPage1
from .page2 import IspdbManager as WizardPage2
from .pages import IspdbManager as WizardPages
from .page5 import IspdbManager as WizardPage5

import traceback


class IspdbController(unohelper.Base,
                      XWizardController):
    def __init__(self, ctx, wizard, sender):
        self._ctx = ctx
        self._wizard = wizard
        self._model = IspdbModel(ctx, sender)

    @property
    def Model(self):
        return self._model

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
                page = WizardPages(self._ctx, self._wizard, self._model, pageid, parent, SMTP)
            elif pageid == 4:
                page = WizardPages(self._ctx, self._wizard, self._model, pageid, parent, IMAP)
            elif pageid == 5:
                page = WizardPage5(self._ctx, self._wizard, self._model, pageid, parent)
            msg += " Done"
            logMessage(self._ctx, INFO, msg, 'IspdbWizard', 'createPage()')
            return page
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.format_exc())
            print(msg)

    def getPageTitle(self, pageid):
        return self._model.getPageStep(pageid)

    def canAdvance(self):
        return True

    def onActivatePage(self, pageid):
        msg = "PageId: %s..." % pageid
        title = self._model.getPageTitle(pageid)
        self._wizard.setTitle(title)
        msg += " Done"
        logMessage(self._ctx, INFO, msg, 'WizardController', 'onActivatePage()')

    def onDeactivatePage(self, pageid):
        pass

    def confirmFinish(self):
        return True
