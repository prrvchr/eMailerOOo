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
                 PropertySet,
                 XWizardPage):
    def __init__(self, ctx, pageid, window, manager):
        msg = "PageId: %s loading ..." % pageid
        self.ctx = ctx
        self.PageId = pageid
        self.Window = window
        self._manager = manager
        self._manager.initPage(pageid, window)
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardPage', '__init__()')

    # XWizardPage
    def activatePage(self):
        try:
            msg = "PageId: %s ..." % self.PageId
            if self.PageId == 2:
                self._manager.activatePage2(self.Window, self._updateProgress)
            elif self.PageId == 3:
                self._manager.activatePage3(self.Window)
            msg += " Done"
            logMessage(self.ctx, INFO, msg, 'WizardPage', 'activatePage()')
        except Exception as e:
            msg = "WizardPage.activatePage() Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def commitPage(self, reason):
        try:
            msg = "PageId: %s ..." % self.PageId
            forward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.FORWARD')
            backward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.BACKWARD')
            finish = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.FINISH')
            if self.PageId == 1:
                self._manager.commitPage1(self.Window)
            elif self.PageId == 2:
                self._manager.commitPage2()
            elif self.PageId == 3:
                pass
            msg += " Done"
            logMessage(self.ctx, INFO, msg, 'WizardPage', 'commitPage()')
            return True
        except Exception as e:
            msg = "WizardPage.commitPage() Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def canAdvance(self):
        return self._manager.canAdvancePage(self.PageId, self.Window)

    def _updateProgress(self, value, offset=0):
        self._manager.updateProgress(self.Window, value, offset)

    def _getPropertySetInfo(self):
        properties = {}
        readonly = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.READONLY')
        properties['PageId'] = getProperty('PageId', 'short', readonly)
        properties['Window'] = getProperty('Window', 'com.sun.star.awt.XWindow', readonly)
        return properties
