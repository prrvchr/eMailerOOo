#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.awt import XContainerWindowEventHandler

from unolib import PropertySet
from unolib import createService
from unolib import getProperty
from unolib import getStringResource
from unolib import getPropertyValueSet

from .dbtools import getValueFromResult
from .dbqueries import getSqlQuery
from .wizardtools import getRowSetOrders
from .wizardtools import getSelectedItems
from .wizardtools import getStringItemList

from .configuration import g_identifier
from .configuration import g_extension

from .logger import logMessage

import traceback


class WizardHandler(unohelper.Base,
                    XContainerWindowEventHandler):
    def __init__(self, ctx, wizard):
        self.ctx = ctx
        self._wizard = wizard

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        try:
            handled = False
            if method == 'TextChange':
                handled = self._updateUI(window, event)
            return handled
        except Exception as e:
            print("WizardHandler.callHandlerMethod() ERROR: %s - %s" % (e, traceback.print_exc()))
    def getSupportedMethodNames(self):
        return ('TextChange',)

    def _updateUI(self, window, event):
        try:
            control = event.Source
            tag = control.Model.Tag
            handled = False
            if tag == 'EmailAddress':
                self._wizard.updateTravelUI()
                handled = True
            return handled
        except Exception as e:
            print("WizardHandler._updateUI() ERROR: %s - %s" % (e, traceback.print_exc()))
