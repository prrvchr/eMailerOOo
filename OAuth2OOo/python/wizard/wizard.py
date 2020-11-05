#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizard
from com.sun.star.lang import XInitialization

from com.sun.star.lang import IllegalArgumentException
from com.sun.star.util import InvalidStateException
from com.sun.star.container import NoSuchElementException

from unolib import getDialog
from unolib import getInterfaceTypes

from .wizardmanager import WizardManager

from .configuration import g_extension

from .logger import getMessage
g_message = 'wizard'

import traceback


class Wizard(unohelper.Base,
             XWizard,
             XInitialization):
    def __init__(self, ctx, auto=-1, resize=False, parent=None):
        try:
            self.ctx = ctx
            self._helpUrl = ''
            self._manager = WizardManager(self.ctx, auto, resize, parent)
            self._manager.initWizard()
            print("Wizard.__init__()")
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    @property
    def HelpURL(self):
        return self._helpUrl
    @HelpURL.setter
    def HelpURL(self, url):
        self._helpUrl = url
        self.DialogWindow.getControl('CommandButton5').Model.Enabled = url != ''
    @property
    def DialogWindow(self):
        return self._manager._view.DialogWindow

    # XInitialization
    def initialize(self, arguments):
        if not isinstance(arguments, tuple) or len(arguments) != 2:
            raise self._getIllegalArgumentException(0, 101)
        paths, controller = arguments
        if not isinstance(paths, tuple) or len(paths) < 1:
            raise self._getIllegalArgumentException(0, 102)
        unotype = uno.getTypeByName('com.sun.star.ui.dialogs.XWizardController')
        if unotype not in getInterfaceTypes(controller):
            raise self._getIllegalArgumentException(0, 103)
        self._manager.setPaths(paths)
        self._manager._controller = controller

    # XWizard
    def getCurrentPage(self):
        return self._manager.getCurrentPage()

    def enableButton(self, button, enabled):
        self._manager.enableButton(button, enabled)

    def setDefaultButton(self, button):
        self._manager.setDefaultButton(button)

    def travelNext(self):
        return self._manager.travelNext()

    def travelPrevious(self):
        return self._manager.travelPrevious()

    def enablePage(self, pageid, enabled):
        if not self._manager.isPathInitialized():
            raise self._getInvalidStateException(111)
        path = self._manager.getCurrentPath()
        if pageid not in path:
            raise self._getNoSuchElementException(112)
        if pageid == self._manager._model.getCurrentPageId():
            raise self._getInvalidStateException(113)
        self._manager.enablePage(pageid, enabled)

    def updateTravelUI(self):
        self._manager.updateTravelUI()

    def advanceTo(self, pageid):
        return self._manager.advanceTo(pageid)

    def goBackTo(self, pageid):
        return self._manager.goBackTo(pageid)

    def activatePath(self, index, final):
        if not self._manager.isMultiPaths():
            return
        if index not in range(self._manager.getPathsLength()):
            raise self._getNoSuchElementException(121)
        path = self._manager.getPath(index)
        page = self._manager._model.getCurrentPageId()
        if page != -1 and page not in path:
            raise self._getInvalidStateException(122)
        self._manager.activatePath(index, final)

    # XExecutableDialog -> XWizard
    def setTitle(self, title):
        self.DialogWindow.setTitle(title)

    def execute(self):
        return self._manager.executeWizard(self.DialogWindow)

    # Private methods
    def _getIllegalArgumentException(self, position, code):
        e = IllegalArgumentException()
        e.ArgumentPosition = position
        e.Message = getMessage(self.ctx, g_message, code)
        e.Context = self
        return e

    def _getInvalidStateException(self, code):
        e = InvalidStateException()
        e.Message = getMessage(self.ctx, g_message, code)
        e.Context = self
        return e

    def _getNoSuchElementException(self, code):
        e = NoSuchElementException()
        e.Message = getMessage(self.ctx, g_message, code)
        e.Context = self
        return e
