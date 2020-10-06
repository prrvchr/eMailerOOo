#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizardPage

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import PropertySet
from unolib import getProperty
from unolib import getStringResource

from .logger import logMessage

from .configuration import g_identifier
from .configuration import g_extension

import traceback


class WizardPage(unohelper.Base,
                 PropertySet,
                 XWizardPage):
    def __init__(self, ctx, id, window, handler):
        msg = "PageId: %s loading ..." % id
        self.ctx = ctx
        self.PageId = id
        self.Window = window
        self._handler = handler
        if id == 1:
            self._initPage1()
        elif id == 2:
            self._initPage2()
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardPage', '__init__()')

    def _initPage1(self):
        pass

    def _initPage2(self):
        pass

    # XWizardPage
    def activatePage(self):
        msg = "PageId: %s ..." % self.PageId
        if self.PageId == 1:
            pass
        elif self.PageId == 2:
            pass
        elif self.PageId == 3:
            pass
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardPage', 'activatePage()')

    def commitPage(self, reason):
        msg = "PageId: %s ..." % self.PageId
        forward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.FORWARD')
        backward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.BACKWARD')
        finish = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.FINISH')
        if self.PageId == 1:
            pass
        elif self.PageId == 2:
            pass
        elif self.PageId == 3:
            pass
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardPage', 'commitPage()')
        return True

    def canAdvance(self):
        advance = False
        if self.PageId == 1:
            advance = self.Window.getControl("TextField1").Text != ''
        elif self.PageId == 2:
            advance = True
        elif self.PageId == 3:
            pass
        return advance

    def _getPropertySetInfo(self):
        properties = {}
        readonly = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.READONLY')
        properties['PageId'] = getProperty('PageId', 'short', readonly)
        properties['Window'] = getProperty('Window', 'com.sun.star.awt.XWindow', readonly)
        return properties
