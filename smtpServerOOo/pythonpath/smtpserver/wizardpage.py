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
        self._manager = manager
        self._manager.initPage(pageid, window)
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardPage', '__init__()')

    @property
    def Window(self):
        return self._manager.getView(self.PageId).Window

    # XWizardPage
    def activatePage(self):
        try:
            msg = "PageId: %s ..." % self.PageId
            self._manager.activatePage(self.PageId)
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
            self._manager.commitPage(self.PageId, reason)
            msg += " Done"
            logMessage(self.ctx, INFO, msg, 'WizardPage', 'commitPage()')
            return True
        except Exception as e:
            msg = "WizardPage.commitPage() Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def canAdvance(self):
        print("WizardPage.canAdvance() %s" % self.PageId)
        return self._manager.canAdvancePage(self.PageId)

    def _getPropertySetInfo(self):
        properties = {}
        readonly = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.READONLY')
        properties['PageId'] = getProperty('PageId', 'short', readonly)
        properties['Window'] = getProperty('Window', 'com.sun.star.awt.XWindow', readonly)
        return properties
