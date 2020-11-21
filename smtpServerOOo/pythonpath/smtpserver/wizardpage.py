#!
# -*- coding: utf_8 -*-

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
