#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs.WizardButton import NEXT
from com.sun.star.ui.dialogs.WizardButton import PREVIOUS
from com.sun.star.ui.dialogs.WizardButton import FINISH
from com.sun.star.ui.dialogs.WizardButton import CANCEL
from com.sun.star.ui.dialogs.WizardButton import HELP

from unolib import getDialog

from .configuration import g_extension

import traceback


class WizardView(unohelper.Base):
    def __init__(self, ctx, handler, xdl, parent):
        self.ctx = ctx
        self._spacer = 5
        self._roadmap = 'RoadmapControl1'
        self._point = uno.createUnoStruct('com.sun.star.awt.Point', 0, 0)
        self._size = uno.createUnoStruct('com.sun.star.awt.Size', 85, 180)
        self._button = {CANCEL: 1, FINISH: 2, NEXT: 3, PREVIOUS: 4, HELP: 5}
        self.DialogWindow = getDialog(self.ctx, g_extension, xdl, handler, parent)

    def getRoadmapPosition(self):
        return self._point

    def getRoadmapSize(self):
        return self._size

    def getRoadmapName(self):
        return self._roadmap

    def getRoadmapTitle(self):
        return self._getRoadmapTitle()

    def initWizard(self, roadmap):
        self.DialogWindow.Model.insertByName(roadmap.Name, roadmap)
        return self._getRoadmap()

    def getPageStep(self, model, pageid):
        return model.resolveString(self._getRoadmapStep(pageid))

# WizardView setter methods
    def setDialogSize(self, page):
        button = self._getButton(HELP).Model
        button.PositionY  = page.Height + self._spacer
        dialog = self.DialogWindow.Model
        dialog.Height = button.PositionY + button.Height + self._spacer
        dialog.Width = page.PositionX + page.Width
        # We assume all buttons are named appropriately
        for i in (1,2,3,4):
            self._setButtonPosition(i, button.PositionY, dialog.Width)

    def enableButton(self, button, enabled):
        self._getButton(button).Model.Enabled = enabled

    def setDefaultButton(self, button):
        self._getButton(button).Model.DefaultButton = True

    def updateButtonPrevious(self, enabled):
        self._getButton(PREVIOUS).Model.Enabled = enabled

    def updateButtonNext(self, enabled):
        button = self._getButton(NEXT).Model
        button.Enabled = enabled
        if enabled:
            button.DefaultButton = True

    def updateButtonFinish(self, enabled):
        button = self._getButton(FINISH).Model
        button.Enabled = enabled
        if enabled:
            button.DefaultButton = True

# WizardView private methods
    def _setButtonPosition(self, step, y, width):
        # We assume that all buttons are the same Width
        button = self._getButtonByIndex(step).Model
        button.PositionX = width - step * (button.Width + self._spacer)
        button.PositionY = y

    def _getButton(self, button):
        index = self._button.get(button)
        return self._getButtonByIndex(index)

# WizardView private message methods
    def _getRoadmapTitle(self):
        return 'Wizard.Roadmap.Text'

    def _getRoadmapStep(self, pageid):
        return 'PageWizard%s.Step' % pageid

# WizardView private control methods
    def _getRoadmap(self):
        return self.DialogWindow.getControl(self.getRoadmapName())

    def _getButtonByIndex(self, index):
        return self.DialogWindow.getControl('CommandButton%s' % index)
