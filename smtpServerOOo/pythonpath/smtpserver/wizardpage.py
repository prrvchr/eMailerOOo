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
    def __init__(self, ctx, pageid, window, handler):
        msg = "PageId: %s loading ..." % pageid
        self.ctx = ctx
        self.PageId = pageid
        self.Window = window
        self._handler = handler
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)
        if pageid == 1:
            print("WizardPage.__init__() %s" % handler._model.Email)
            window.getControl('TextField1').Text = handler.getEmail()
        elif pageid == 2:
            pass
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardPage', '__init__()')

    # XWizardPage
    def activatePage(self):
        msg = "PageId: %s ..." % self.PageId
        if self.PageId == 2:
            email = self._handler.getEmail()
            self._setPageTitle(email)
            self._updateProgress(0)
            self._handler._model._datasource.loadSmtpConfig(email, self._updateProgress)
        elif self.PageId == 3:
            self._setPageTitle(self._handler.getEmail())
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
        return self._handler.canAdvancePage(self.PageId)

    def _setPageTitle(self, title):
        text = self._stringResource.resolveString('PageWizard%s.Label1.Label' % self.PageId)
        self.Window.getControl('Label1').Text = text % title

    def _updateProgress(self, value):
        self.Window.getControl('ProgressBar1').Value = value
        text = self._stringResource.resolveString('PageWizard2.Label2.Label.%s' % value)
        self.Window.getControl('Label2').Text = text

    def _getPropertySetInfo(self):
        properties = {}
        readonly = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.READONLY')
        properties['PageId'] = getProperty('PageId', 'short', readonly)
        properties['Window'] = getProperty('Window', 'com.sun.star.awt.XWindow', readonly)
        return properties
