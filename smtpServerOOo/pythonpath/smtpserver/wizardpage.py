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

from com.sun.star.ui.dialogs import XWizardPage

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import PropertySet
from unolib import getProperty

from .logger import logMessage

import traceback


class WizardPage(unohelper.Base,
                 XWizardPage):
    def __init__(self, ctx, pageid, window, manager):
        self.ctx = ctx
        self.PageId = pageid
        self._manager = manager
        self._manager.initPage(pageid, window)

    @property
    def Window(self):
        return self._manager.getView(self.PageId).Window

    # XWizardPage
    def activatePage(self):
        try:
            self._manager.activatePage(self.PageId)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            logMessage(self.ctx, SEVERE, msg, 'WizardPage', 'activatePage()')
            print("WizardPage.activatePage() %s" % msg)

    def commitPage(self, reason):
        try:
            return self._manager.commitPage(self.PageId, reason)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            logMessage(self.ctx, SEVERE, msg, 'WizardPage', 'commitPage()')
            print("WizardPage.commitPage() %s" % msg)

    def canAdvance(self):
        return self._manager.canAdvancePage(self.PageId)
