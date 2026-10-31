#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-25 https://prrvchr.github.io                                  ║
║                                                                                    ║
║   Permission is hereby granted, free of charge, to any person obtaining            ║
║   a copy of this software and associated documentation files (the "Software"),     ║
║   to deal in the Software without restriction, including without limitation        ║
║   the rights to use, copy, modify, merge, publish, distribute, sublicense,         ║
║   and/or sell copies of the Software, and to permit persons to whom the Software   ║
║   is furnished to do so, subject to the following conditions:                      ║
║                                                                                    ║
║   The above copyright notice and this permission notice shall be included in       ║
║   all copies or substantial portions of the Software.                              ║
║                                                                                    ║
║   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,                  ║
║   EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES                  ║
║   OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.        ║
║   IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY             ║
║   CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,             ║
║   TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE       ║
║   OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.                                    ║
║                                                                                    ║
╚════════════════════════════════════════════════════════════════════════════════════╝
"""

import uno

from ..unotool import executeDesktopDispatch
from ..unotool import getConfiguration
from ..unotool import getStringResource
from ..unotool import saveWindowPosition

from ..helper import getMailSpooler

from ..configuration import g_identifier

import traceback


class ControllerModel():
    def __init__(self, ctx):
        self._ctx = ctx
        self._spooler = None
        self._indicator = None
        self._config = getConfiguration(ctx, g_identifier, True)
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'SpoolerController')
        self._resources = {'Title':       'SpoolerController.Title',
                           'Button':      'SpoolerController.CommandButton1.Label.%s',
                           'Label':       'SpoolerController.Label1.Label.%s'}

# ControllerModel getter methods
    def getViewData(self):
        return self._getDialogPosition(), self._getDialogTitle()

    def getLabelState(self, state):
        return self._getLabelState(state)

    def initToggleState(self):
        state = int(self._spooler.isRunning())
        return state, self._getToggleState(state), self._getLabelState(state)

    def getToggleState(self, state, reset=False):
        if reset:
            self._indicator = None
        return state, self._getToggleState(state), self._getLabelState(state)

    def hasIndicator(self):
        return self._indicator is not None

    def getSpooler(self):
        return self._spooler

# ControllerModel setter methods
    def initSpooler(self, listener):
        print("ControllerModel.initSpooler() 1")
        self._spooler = getMailSpooler(self._ctx)
        self._spooler.addSpoolerListener(listener)
        print("ControllerModel.initSpooler() 2")

    def startSpooler(self, indicator):
        self._spooler.start(indicator)
        self._indicator = indicator

    def stopSpooler(self):
        self._spooler.stop()
        self._indicator = None

    def addSpoolerIndicator(self, indicator):
        if self._spooler:
            self._spooler.addStatusIndicator(indicator)
            self._indicator = indicator

    def removeSpoolerIndicator(self, indicator):
        if self._spooler:
            self._spooler.removeStatusIndicator(indicator)
            self._indicator = None

    def disposeSpooler(self, listener):
        print("ControllerModel.disposeSpooler() 1")
        if self._spooler:
            self._spooler.removeSpoolerListener(listener)
            if self._indicator:
                self._spooler.removeStatusIndicator(self._indicator)
                print("ControllerModel.disposeSpooler() 2")
            self._spooler.dispose()
        self._spooler = None
        self._indicator = None
        print("ControllerModel.disposeSpooler() 3")

    def savePosition(self, position):
        saveWindowPosition(self._config, position, 'ControllerPosition')

# ControllerModel private getter methods
    def _getDialogPosition(self):
        return uno.createUnoStruct('com.sun.star.awt.Point', *self._config.getByName('ControllerPosition'))

# ControllerModel private setter methods
    def _getDialogTitle(self):
        resource = self._resources.get('Title')
        return self._resolver.resolveString(resource)

    def _getLabelState(self, state):
        resource = self._resources.get('Label') % state
        return self._resolver.resolveString(resource)

    def _getToggleState(self, state):
        resource = self._resources.get('Button') % state
        return self._resolver.resolveString(resource)
