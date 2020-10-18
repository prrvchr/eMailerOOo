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

import validators
import traceback


class WizardHandler(unohelper.Base,
                    XContainerWindowEventHandler):
    def __init__(self, ctx, wizard, model):
        self.ctx = ctx
        self._wizard = wizard
        self._model = model

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        try:
            handled = False
            if method == 'TextChange':
                handled = self._updateUI(window, event)
            elif method == 'ChangeAuthentication':
                self._changeAuthentication(window, event.Source)
                handled = True
            elif method == 'SmtpConnect':
                self._smtpConnect(window, event)
                handled = True
            return handled
        except Exception as e:
            print("WizardHandler.callHandlerMethod() ERROR: %s - %s" % (e, traceback.print_exc()))
    def getSupportedMethodNames(self):
        return ('TextChange', 'ChangeAuthentication', 'SmtpConnect')

    def getEmail(self):
        return self._model.Email

    def setEmail(self, email):
        self._model.Email = email

    def _isEmailValid(self):
        if validators.email(self.getEmail()):
           return True
        return False

    def canAdvancePage(self, page):
        advance = False
        if page == 1:
            advance = self._isEmailValid()
        elif page == 2:
           advance = True
        elif page == 2:
           advance = True
        return advance

    def getDomain(self):
        return self.getEmail().split('@').pop()

    def _updateUI(self, window, event):
        try:
            control = event.Source
            tag = control.Model.Tag
            handled = False
            if tag == 'EmailAddress':
                self.setEmail(control.Text)
                self._wizard.updateTravelUI()
                handled = True
            return handled
        except Exception as e:
            print("WizardHandler._updateUI() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _changeAuthentication(self, window, control):
        index = control.getSelectedItemPos()
        if index == 0:
            window.getControl('Label6').Model.Enabled = False
            window.getControl('TextField2').Model.Enabled = False
            window.getControl('Label7').Model.Enabled = False
            window.getControl('TextField3').Model.Enabled = False
        elif index == 3:
            window.getControl('Label6').Model.Enabled = True
            window.getControl('TextField2').Model.Enabled = True
            window.getControl('Label7').Model.Enabled = False
            window.getControl('TextField3').Model.Enabled = False
        else:
            window.getControl('Label6').Model.Enabled = True
            window.getControl('TextField2').Model.Enabled = True
            window.getControl('Label7').Model.Enabled = True
            window.getControl('TextField3').Model.Enabled = True
        print("WizardHandler._changeAuthentication()")

    def _smtpConnect(self, window, event):
        print("WizardHandler._smtpConnect()")
