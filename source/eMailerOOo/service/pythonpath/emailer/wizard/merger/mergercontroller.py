#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-25 https://prrvchr.github.io                                  ║
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

from com.sun.star.lang import XComponent

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.ui.dialogs import XWizardController

from com.sun.star.util import CloseVetoException
from com.sun.star.util import XCloseable

from .mergermodel import MergerModel

from .page1 import MergerManager as WizardPage1
from .page2 import MergerManager as WizardPage2
from .page3 import MergerManager as WizardPage3

from ...unotool import getStringResource
from ...unotool import getTopWindowPosition

from ...configuration import g_identifier

import traceback


class MergerController(unohelper.Base,
                       XWizardController,
                       XComponent,
                       XCloseable):
    def __init__(self, ctx, wizard, datasource, document):
        self._ctx = ctx
        self._listeners = []
        self._wizard = wizard
        self._model = MergerModel(ctx, datasource, document)
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'MergerController')

# XCloseable
    def close(self, ownership):
        if not ownership:
            if self._model.isClosing():
                raise CloseVetoException()
            self._model.setClosing()
            self._close()
            if self._model.hasDispatch():
                self._model.cancelDispatch()
                raise CloseVetoException()

    def addCloseListener(self, listener):
        if listener not in self._listeners:
            self._listeners.append(listener)

    def removeCloseListener(self, listener):
        if listener in self._listeners:
            self._listeners.remove(listener)

# XComponent
    def dispose(self):
        self._model.dispose()

    def addEventListener(self, listener):
        pass

    def removeEventListener(self, listener):
        pass

# XWizardController
    def createPage(self, parent, pageid):
        try:
            if pageid == 1:
                page = WizardPage1(self._ctx, self._wizard, self._model, pageid, parent)
            elif pageid == 2:
                page = WizardPage2(self._ctx, self._wizard, self._model, pageid, parent)
            elif pageid == 3:
                page = WizardPage3(self._ctx, self._wizard, self._model, pageid, parent)
            return page
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.format_exc())
            print(msg)

    def getPageTitle(self, pageid):
        return self._model.getPageStep(self._resolver, pageid)

    def canAdvance(self):
        return True

    def onActivatePage(self, pageid):
        title = self._model.getPageTitle(self._resolver, pageid)
        self._wizard.setTitle(title)

    def onDeactivatePage(self, pageid):
        pass

    def confirmFinish(self):
        return True

    def _close(self):
        position = getTopWindowPosition(self._wizard.DialogWindow)
        self._model.saveView(position)

