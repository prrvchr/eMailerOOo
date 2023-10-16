#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020 https://prrvchr.github.io                                     ║
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

import unohelper

from com.sun.star.awt.FontWeight import NORMAL

from .ispdbview import IspdbView
from .ispdbhandler import WindowHandler

import traceback


class IspdbManager(unohelper.Base):
    def __init__(self, ctx, wizard, model, pageid, parent):
        self._ctx = ctx
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._view = IspdbView(ctx, WindowHandler(self), parent)
        self._finish = False

# XWizardPage
    @property
    def PageId(self):
        return self._pageid
    @property
    def Window(self):
        return self._view.getWindow()

    def activatePage(self):
        self._finish = False
        self._model.Offline = 0
        self._wizard.activatePath(1, False)
        self._wizard.enablePage(1, False)
        label = self._model.getPageLabel(self._pageid)
        self._view.setPageLabel(label % self._model.Email)
        self._wizard.updateTravelUI()
        self._model.getServerConfig(self.updateProgress, self.updateView)

    def commitPage(self, reason):
        self._finish = False
        return True

    def canAdvance(self):
        return self._finish

# IspdbManager setter methods
    def enableIMAP(self, imap):
        self._activatePath(self._model.setUserConfig(imap), imap)

    def updateProgress(self, value, offset=0, style=NORMAL):
        if not self._model.isDisposed():
            message = self._model.getProgressMessage(value + offset)
            self._view.updateProgress(value, message, style)

    def updateView(self, user, config, imap, offline, auto):
        if not self._model.isDisposed():
            self._model.setServerConfig(user, config, imap, offline)
            title = self._model.getPageTitle(self._pageid)
            self._wizard.setTitle(title)
            self._wizard.enablePage(1, True)
            if not auto:
                self._view.enableIMAP(imap)
            self._activatePath(offline, imap)
            self._finish = True
            self._wizard.updateTravelUI()

    def _activatePath(self, offline, imap):
        path = offline * 2 + imap
        self._wizard.activatePath(path, True)
