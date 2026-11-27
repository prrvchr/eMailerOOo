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

import unohelper

from com.sun.star.ucb.ConnectionMode import OFFLINE

from .ispdbview import IspdbView
from .ispdbhandler import WindowHandler

from ...unotool import getStringResource

from ...configuration import g_identifier


class IspdbManager(unohelper.Base):
    def __init__(self, ctx, wizard, model, pageid, parent):
        self._wizard = wizard
        self._model = model
        self._pageid = pageid
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'IspdbPage2')
        self._view = IspdbView(ctx, WindowHandler(self), parent)
        self._finished = False

# XWizardPage
    @property
    def PageId(self):
        return self._pageid

    @property
    def Window(self):
        return self._view.getWindow()

    def activatePage(self):
        self._finished = False
        self._model.Offline = OFFLINE
        self._wizard.activatePath(1, False)
        self._wizard.enablePage(1, False)
        self._view.initSearch(self._model.getPageLabel(self._resolver, self._pageid, self._model.Email))
        self._wizard.updateTravelUI()
        self._model.getServerConfig(self.onServerConfig, self._view.getIndicator())

    def commitPage(self, _reason):
        if self._view.useReplyTo():
            self._model.setReplyToAddress(self._view.getReplyTo())
        return True

    def canAdvance(self):
        return self._finished and self._model.isEmailValid(self._view.getReplyTo())

# IspdbManager setter methods
    def onServerConfig(self, _success, servers, user, mode, _error):
        self._model.Offline = mode
        self._model.setServerConfig(servers)
        self._updateView(user, servers.hasConfig() and user.isNew())
        self._updatePath(mode)

    def enableReplyTo(self, state):
        self._view.setReplyToAddress(state, self._model.enableReplyTo(state))

    def changeReplyTo(self):
        self._wizard.updateTravelUI()

    def enableIMAP(self, imap):
        self._activatePath(imap, self._model.enableIMAP(imap))

# IspdbManager private methods
    def _updateView(self, user, auto):
        if not self._model.isDisposed():
            self._wizard.setTitle(self._model.getPageTitle(self._resolver, 2))
            self._wizard.enablePage(1, True)
            self._view.commitSearch(user.getReplyToState(), user.getImapState(), user.ReplyToAddress, auto)

    def _updatePath(self, mode):
        if not self._model.isDisposed():
            self._activatePath(self._view.getIMAP(), mode)
            self._finished = True
            self._wizard.updateTravelUI()

    def _activatePath(self, imap, mode):
        path = mode * 2 + imap
        self._wizard.activatePath(path, True)
