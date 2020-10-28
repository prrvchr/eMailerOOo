#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs.WizardButton import NEXT
from com.sun.star.ui.dialogs.WizardButton import PREVIOUS
from com.sun.star.ui.dialogs.WizardButton import FINISH
from com.sun.star.ui.dialogs.WizardButton import CANCEL
from com.sun.star.ui.dialogs.WizardButton import HELP

import traceback


class WizardView(unohelper.Base):
    def __init__(self, ctx):
        self.ctx = ctx
        self._spacer = 5
        self._roadmap = 'RoadmapControl1'
        self._point = uno.createUnoStruct('com.sun.star.awt.Point', 0, 0)
        self._size = uno.createUnoStruct('com.sun.star.awt.Size', 85, 180)
        self._button = {CANCEL: 1, FINISH: 2, NEXT: 3, PREVIOUS: 4, HELP: 5}

    def getRoadmapPosition(self):
        return self._point

    def getRoadmapSize(self):
        return self._size

    def getRoadmapName(self):
        return self._roadmap

    def getRoadmapTitle(self):
        return self._getRoadmapTitle()

    def initWizard(self, window, roadmap):
        window.getModel().insertByName(roadmap.Name, roadmap)
        return self._getRoadmap(window)

# WizardView setter methods
    def setDialogStep(self, window, step):
        window.getModel().Step = step

    def setDialogSize(self, window, page):
        button = self._getButton(window, HELP).getModel()
        button.PositionY  = page.Height + self._spacer
        dialog = window.getModel()
        dialog.Height = button.PositionY + button.Height + self._spacer
        dialog.Width = page.PositionX + page.Width
        # We assume all buttons are named appropriately
        for i in (1,2,3,4):
            self._setButtonPosition(window, i, button.PositionY, dialog.Width)

    def enableButton(self, window, button, enabled):
        self._getButton(window, button).Model.Enabled = enabled

    def setDefaultButton(self, window, button):
        self._getButton(window, button).Model.DefaultButton = True

    def updateButtonPrevious(self, window, enabled):
        self._getButton(window, PREVIOUS).Model.Enabled = enabled

    def updateButtonNext(self, window, enabled):
        button = self._getButton(window, NEXT).Model
        button.Enabled = enabled
        if enabled:
            button.DefaultButton = True

    def updateButtonFinish(self, window, enabled):
        button = self._getButton(window, FINISH).Model
        button.Enabled = enabled
        if enabled:
            button.DefaultButton = True

# WizardView getter methods
    def getDialogStep(self, window):
        return window.getModel().Step

# WizardView private methods
    def _setButtonPosition(self, window, step, y, width):
        # We assume that all buttons are the same Width
        button = window.getControl(self._getButtonName(step)).getModel()
        button.PositionX = width - step * (button.Width + self._spacer)
        button.PositionY = y

    def _getButton(self, window, button):
        return window.getControl(self._getButtonName(self._button.get(button)))

# WizardView private message methods
    def _getRoadmapTitle(self):
        return 'Wizard.Roadmap.Text'

# WizardView private control methods
    def _getRoadmap(self, window):
        return window.getControl(self._roadmap)

    def _getButtonName(self, index):
        return 'CommandButton%s' % index
