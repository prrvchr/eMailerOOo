#!
# -*- coding: utf_8 -*-

import uno
import unohelper

import traceback


class WizardView(unohelper.Base):
    def __init__(self, ctx):
        self.ctx = ctx
        self._spacer = 5
        self._roadmap = 'RoadmapControl1'
        self._point = uno.createUnoStruct('com.sun.star.awt.Point', 0, 0)
        self._size = uno.createUnoStruct('com.sun.star.awt.Size', 85, 180)

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
        button = self._getButtonHelp(window).getModel()
        button.PositionY  = page.Height + self._spacer
        dialog = window.getModel()
        dialog.Height = button.PositionY + button.Height + self._spacer
        dialog.Width = page.PositionX + page.Width
        # We assume all buttons are named appropriately
        for i in (1,2,3,4):
            self._setButtonPosition(window, i, button.PositionY, dialog.Width)

    def enableButtonHelp(self, window, enabled)
        self._getButtonHelp(window).Model.Enabled = enabled

    def enableButtonPrevious(self, window, enabled)
        self._getButtonPrevious(window).Model.Enabled = enabled

    def enableButtonNext(self, window, enabled)
        self._getButtonNext(window).Model.Enabled = enabled

    def enableButtonFinish(self, window, enabled)
        self._getButtonFinish(window).Model.Enabled = enabled

    def enableButtonCancel(self, window, enabled)
        self._getButtonCancel(window).Model.Enabled = enabled

    def setDefaultButtonHelp(self, window)
        self._getButtonHelp(window).Model.DefaultButton = True

    def setDefaultButtonPrevious(self, window)
        self._getButtonPrevious(window).Model.DefaultButton = True

    def setDefaultButtonNext(self, window)
        self._getButtonNext(window).Model.DefaultButton = True

    def setDefaultButtonFinish(self, window)
        self._getButtonFinish(window).Model.DefaultButton = True

    def setDefaultButtonCancel(self, window)
        self._getButtonCancel(window).Model.DefaultButton = True

    def updateButtonPrevious(self, window, enabled):
        self._getButtonPrevious(window).Model.Enabled = enabled

    def updateButtonNext(self, window, enabled):
        button = self._getButtonNext(window).Model
        button.Enabled = enabled
        if enabled:
            button.DefaultButton = True

    def updateButtonFinish(self, window, enabled):
        button = self._getButtonFinish(window).Model
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

# WizardView private message methods
    def _getRoadmapTitle(self):
        return 'Wizard.Roadmap.Text'

# WizardView private control methods
    def _getRoadmap(self, window):
        return window.getControl(self._roadmap)

    def _getButtonName(self, index):
        return 'CommandButton%s' % index

    def _getButtonHelp(self, window):
        return window.getControl(self._getButtonName(5))

    def _getButtonPrevious(self, window):
        return window.getControl(self._getButtonName(4))

    def _getButtonNext(self, window):
        return window.getControl(self._getButtonName(3))

    def _getButtonFinish(self, window):
        return window.getControl(self._getButtonName(2))

    def _getButtonCancel(self, window):
        return window.getControl(self._getButtonName(1))

