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

from com.sun.star.uno import Exception as UnoException

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .pagemodel import PageModel
from .pageview import PageView

from unolib import createService

from .logger import logMessage
from .logger import getMessage
g_message = 'pagemanager'


import traceback


class PageManager(unohelper.Base):
    def __init__(self, ctx, wizard, model=None):
        self.ctx = ctx
        self._wizard = wizard
        self._model = PageModel(self.ctx) if model is None else model
        self._views = {}
        print("PageManager.__init__() 1")

    @property
    def View(self):
        pageid = self.Wizard.getCurrentPage().PageId
        return self.getView(pageid)
    @property
    def Model(self):
        return self._model
    @property
    def Wizard(self):
        return self._wizard

    def getView(self, pageid):
        if pageid in self._views:
            return self._views[pageid]
        print("PageManager.getView ERROR **************************")
        return None

    def getWizard(self):
        return self.Wizard

    def initPage(self, pageid, window):
        view = PageView(self.ctx, window)
        self._views[pageid] = view
        if pageid == 1:
            if view.initPage1(self.Model):
                self.Wizard.updateTravelUI()

    def activatePage(self, pageid):
        if pageid == 1:
            pass

    def canAdvancePage(self, pageid):
        advance = False
        if pageid == 1:
            pass
        elif pageid == 2:
            pass
        elif pageid == 3:
            pass
        elif pageid == 4:
            pass
        return advance

    def commitPage(self, pageid, reason):
        if pageid == 1:
            pass
        elif pageid == 2:
            pass
        elif pageid == 3:
            pass
        elif pageid == 4:
            pass
        return True

    def updateTravelUI(self):
        self.Wizard.updateTravelUI()

    def setPageTitle(self, pageid):
        title = self.getView(pageid).getPageTitle(self.Model, pageid)
        self.Wizard.setTitle(title)
