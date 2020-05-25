#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizard
from com.sun.star.lang import XServiceInfo
from com.sun.star.lang import XInitialization

from unolib import createService
from .configuration import g_extension

import traceback


class Wizard(unohelper.Base,
             XWizard,
             XServiceInfo,
             XInitialization):
    def __init__(self, ctx):
        self.ctx = ctx
        self._wizard = createService(self.ctx, 'com.sun.star.ui.dialogs.Wizard')

    @property
    def HelpURL(self):
        return self._wizard.HelpURL
    @HelpURL.setter
    def HelpURL(self, url):
        self._wizard.HelpURL = url
    @property
    def DialogWindow(self):
        return self._wizard.DialogWindow

    # XInitialization
    def initialize(self, args):
        arguments = ((uno.Any(args[0], args[1]), args[2]), )
        uno.invoke(self._wizard, 'initialize', arguments)

    # XWizard
    def getCurrentPage(self):
        return self._wizard.getCurrentPage()

    def enableButton(self, button, enabled):
        self._wizard.enableButton(button, enabled)

    def setDefaultButton(self, button):
        self._wizard.setDefaultButton(button)

    def travelNext(self):
        print("Wizard.travelNext()")
        return self._wizard.travelNext()

    def travelPrevious(self):
        print("Wizard.travelPrevious()")
        return self._wizard.travelPrevious()

    def enablePage(self, page, enabled):
        self._wizard.enablePage(page, enabled)

    def updateTravelUI(self):
        self._wizard.updateTravelUI()

    def advanceTo(self, page):
        print("Wizard.advanceTo()")
        return self._wizard.advanceTo(page)

    def goBackTo(self, page):
        print("Wizard.goBackTo()")
        return self._wizard.goBackTo(page)

    def activatePath(self, path, final):
        self._wizard.activatePath(path, final)

    # XExecutableDialog
    def setTitle(self, title):
        self._wizard.setTitle(title)
    def execute(self):
        dialog = getDialog(self.ctx, window.Peer, self, g_extension, 'Wizard')
        return self._wizard.execute()

    # XServiceInfo
    def supportsService(self, service):
        return self._wizard.supportsService(service)
    def getImplementationName(self):
        return self._wizard.getImplementationName()
    def getSupportedServiceNames(self):
        return self._wizard.getSupportedServiceNames()
