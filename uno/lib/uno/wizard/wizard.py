#!
# -*- coding: utf_8 -*-

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

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizard

from com.sun.star.lang import XInitialization
from com.sun.star.lang import IllegalArgumentException

from com.sun.star.util import InvalidStateException

from com.sun.star.container import NoSuchElementException

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from .wizardmodel import WizardModel
from .wizardview import WizardView

from ..unotool import hasInterface

from ..logger import getMessage
g_message = 'wizard'

import traceback


class Wizard(unohelper.Base,
             XWizard,
             XInitialization):
    def __init__(self, ctx, auto=-1, resize=False, parent=None):
        try:
            self._ctx = ctx
            self._helpUrl = ''
            self._auto = auto
            self._resize = resize
            self._paths = ()
            self._currentPath = -1
            self._multiPaths = False
            self._controller = None
            self._model = WizardModel(ctx)
            title = self._model.getRoadmapTitle()
            self._view = WizardView(ctx, self, parent, title)
            roadmap = self._view.getRoadmapModel()
            self._model.setRoadmapModel(roadmap)
            print("Wizard.__init__()")
        except Exception as e:
            msg = "Error: %s" % traceback.print_exc()
            print(msg)

# XWizard
    @property
    def HelpURL(self):
        return self._helpUrl
    @HelpURL.setter
    def HelpURL(self, url):
        self._helpUrl = url
        enabled = url != ''
        self._view.enableHelpButton(enabled)

    @property
    def DialogWindow(self):
        return self._view.getDialog()

    # not used?
    def getCurrentPage(self):
        return self._model.getCurrentPage()

    def enableButton(self, button, enabled):
        self._view.enableButton(button, enabled)

    def setDefaultButton(self, button):
        self._view.setDefaultButton(button)

    def travelNext(self):
        page = self._getNextPage()
        if page is not None:
            return self._setCurrentPage(page)
        return False

    def travelPrevious(self):
        page = self._getPreviousPage()
        if page is not None:
            return self._setCurrentPage(page)
        return False

    def enablePage(self, pageid, enabled):
        if not self._isPathInitialized():
            raise self._getInvalidStateException(111)
        path = self._getCurrentPath()
        if pageid not in path:
            raise self._getNoSuchElementException(112)
        if pageid == self._model.getCurrentPageId():
            raise self._getInvalidStateException(113)
        if self._model.enablePage(pageid, enabled):
            self._model.updateRoadmap(self._getFirstPage())

    def updateTravelUI(self):
        self._model.updateRoadmap(self._getFirstPage())
        self._updateButton()

    def advanceTo(self, pageid):
        if page in self._model.getRoadmapPath():
            return self._setCurrentPage(page)
        return False

    def goBackTo(self, pageid):
        if page in self._model.getRoadmapPath():
            return self._setCurrentPage(page)
        return False

    def activatePath(self, index, final):
        if not self._isMultiPaths():
            raise self._getInvalidStateException(121)
        if index not in self._getMultiPathsIndex():
            raise self._getNoSuchElementException(122)
        path = self._paths[index]
        page = self._model.getCurrentPageId()
        if page != -1 and page not in path:
            raise self._getInvalidStateException(123)
        if self._currentPath != index or self._isComplete() != final:
            self._initPath(index, final)
            self._model.updateRoadmap(self._getFirstPage())

# XWizard -> XExecutableDialog
    def setTitle(self, title):
        self._view.setDialogTitle(title)

    def execute(self):
        if not self._isCurrentPathSet():
            self._initPath(0, False)
        self._initPage()
        return self._view.execute()

# XInitialization
    def initialize(self, arguments):
        if not isinstance(arguments, tuple) or len(arguments) != 2:
            raise self._getIllegalArgumentException(0, 101)
        paths, controller = arguments
        if not isinstance(paths, tuple) or len(paths) < 1:
            raise self._getIllegalArgumentException(0, 102)
        interface = 'com.sun.star.ui.dialogs.XWizardController'
        if not hasInterface(controller, interface):
            raise self._getIllegalArgumentException(0, 103)
        self._paths = paths
        self._multiPaths = isinstance(paths[0], tuple)
        self._controller = controller

# Wizard setter methods
    def changeRoadmapStep(self, page):
        pageid = self._model.getCurrentPageId()
        if pageid != page:
            if not self._setCurrentPage(page):
                self._model.setCurrentPageId(pageid)

    def doFinish(self):
        if self._isLastPage():
            reason = self._getCommitReason()
            if self._model.doFinish(reason):
                if self._controller.confirmFinish():
                    self.view.endDialog()

# Wizard private getter methods
    def _isComplete(self):
        return self._model.isComplete()

    def _isLastPage(self):
        return self._isComplete() and self._model.getCurrentPageId() == self._getLastPage()

    def _isFirstPage(self):
        return self._model.getCurrentPageId() == self._getFirstPage()

    def _isCurrentPathSet(self):
        return self._currentPath != -1

    def _getNextPage(self):
        path = self._model.getRoadmapPath()
        page = self._model.getCurrentPageId()
        if page in path:
            i = path.index(page) + 1
            if i < len(path):
                return path[i]
        return None

    def _getPreviousPage(self):
        path = self._model.getRoadmapPath()
        page = self._model.getCurrentPageId()
        if page in path:
            i = path.index(page) - 1
            if i >= 0:
                return path[i]
        return None

    def _setCurrentPage(self, page):
        if self._deactivatePage(page):
            self._setPage(page)
            return True
        return False

    def _deactivatePage(self, new):
        old = self._model.getCurrentPageId()
        reason = self._getCommitReason(old, new)
        if self._model.deactivatePage(old, reason):
            self._controller.onDeactivatePage(old)
            self._model.setPageVisible(old, False)
            return True
        return False

    def _getCommitReason(self, old=None, new=None):
        page = self._model.getCurrentPageId()
        if old is None:
            old = page
        if new is None:
            new = page
        if old < new:
            reason = FORWARD
        elif old > new:
            reason = BACKWARD
        else:
            reason = FINISH
        return reason

    def _setPageModel(self, page):
        model = page.Window.getModel()
        # TODO: Fixed: Resizing should be done, if necessary, instead of modifying the model
        if self._resize:
            self._resizeWizard(model)
            model.PositionY = 0
        else:
            model.PositionX = self._model.getRoadmapWidth()
            model.PositionY = 0
        return model.Name

    def _setPageWindow(self, window, page, name):
        window.addControl(name, page.Window)
        return page

    def _isPathInitialized(self):
        return not self._multiPaths or self._currentPath != -1

    def _getCurrentPath(self):
        if self._isMultiPaths():
            path = self._paths[self._currentPath]
        else:
            path = self._paths
        return path

    def _isMultiPaths(self):
        return self._multiPaths

    def _getMultiPathsIndex(self):
        return range(len(self._paths))

    def _getFirstPage(self):
        return min(self._getActivePath())

    def _getLastPage(self):
        return max(self._getActivePath())

    def _getActivePath(self):
        return self._model.getActivePath(self._getCurrentPath())

    def _isAutoLoad(self, page=None):
        if page is None:
            nextindex = self._getFirstPage()
        else:
            nextindex = self._getCurrentPath().index(page) + 1
        return nextindex < self._auto

    def _getPath(self, index, final):
        if self._isMultiPaths():
            if final:
                paths = self._paths[index]
            else:
                paths = self._getCommunPath(index)
            self._currentPath = index
        else:
            final, paths = True, self._paths
        return final, paths

    def _getCommunPath(self, index):
        paths = []
        i = 0
        pageid = self._model.getCurrentPageId()
        for page in self._paths[index]:
            if page > pageid:
                for j in range(len(self._paths)):
                    if j == index or j == self._currentPath:
                        continue
                    if i >= len(self._paths[j]) or page != self._paths[j][i]:
                        return tuple(paths)
            paths.append(page)
            i += 1
        return tuple(paths)

    def _canAdvance(self):
        return self._controller.canAdvance() and self._model.canAdvance()

    def _initNextPage(self):
        init = False
        page = self._getNextPage()
        if page is not None:
            init = self._setCurrentPage(page) and self._isAutoLoad(page)
        return init

# Wizard private setter methods
    def _initPath(self, index, final):
        complete, paths = self._getPath(index, final)
        self._model.initRoadmap(self._controller, paths, complete)

    def _initPage(self):
        self._setPage(self._getFirstPage())
        nextpage = self._isAutoLoad()
        while nextpage and self._canAdvance():
            nextpage = self._initNextPage()

    def _activatePage(self, page):
        self._controller.onActivatePage(page)
        self._model.activatePage(page)

    def _resizeWizard(self, model):
        self._model.setRoadmapSize(model)
        self._view.setDialogSize(model)
        self._resize = False

    def _setPage(self, pageid):
        if not self._model.hasPage(pageid):
            window = self._view.getDialog()
            page = self._controller.createPage(window.getPeer(), pageid)
            name = self._setPageModel(page)
            self._model.addPage(pageid, self._setPageWindow(window, page, name))
        self._model.setCurrentPageId(pageid)
        self._activatePage(pageid)
        # TODO: Fixed: XWizard.updateTravelUI() must be done after XWizardPage.activatePage()
        self._model.updateRoadmap(self._getFirstPage())
        self._updateButton()
        self._model.setPageVisible(pageid, True)

    def _updateButton(self):
        self._view.updateButtonPrevious(not self._isFirstPage())
        enabled = self._getNextPage() is not None and self._canAdvance()
        self._view.updateButtonNext(enabled)
        enabled = self._isLastPage() and self._canAdvance()
        self._view.updateButtonFinish(enabled)

# Private Exception getter methods
    def _getIllegalArgumentException(self, position, code):
        e = IllegalArgumentException()
        e.ArgumentPosition = position
        e.Message = getMessage(self._ctx, g_message, code)
        e.Context = self
        return e

    def _getInvalidStateException(self, code):
        e = InvalidStateException()
        e.Message = getMessage(self._ctx, g_message, code)
        e.Context = self
        return e

    def _getNoSuchElementException(self, code):
        e = NoSuchElementException()
        e.Message = getMessage(self._ctx, g_message, code)
        e.Context = self
        return e
