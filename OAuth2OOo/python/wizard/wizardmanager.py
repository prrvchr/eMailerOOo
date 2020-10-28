#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from .wizardmodel import WizardModel
from .wizardview import WizardView

import traceback


class WizardManager(unohelper.Base):
    def __init__(self, ctx, auto, resize):
        self.ctx = ctx
        self._auto = auto
        self._resize = resize
        self._paths = ()
        self._currentPath = -1
        self._firstPage = -1
        self._lastPage = -1
        self._multiPaths = False
        self._controller = None
        self._model = WizardModel(self.ctx)
        self._view = WizardView(self.ctx)
        print("WizardManager.__init__()")

    def initWizard(self, window, handler):
        model = self._model.initWizard(window, self._view)
        roadmap = self._view.initWizard(window, model)
        roadmap.addItemListener(handler)

    def getCurrentPage(self):
        return self._model.getCurrentPage()

    def getCurrentPath(self):
        if self.isMultiPaths():
            return self._paths[self._currentPath]
        return self._paths

    def setPaths(self, paths):
        self._paths = paths
        self._multiPaths = isinstance(paths[0], tuple)

    def getPath(self, index):
        return self._paths[index]

    def isMultiPaths(self):
        return self._multiPaths

    def getPathsLength(self):
        return len(self._paths)

    def activatePath(self, index, final):
        if self._currentPath != index or self._model.isFinal() != final:
            self._initPath(index, final)

    def doFinish(self, dialog):
        if self._isLastPage():
            reason = self._getCommitReason()
            if self._model.doFinish(reason):
                if self._controller.confirmFinish():
                    dialog.endDialog(OK)

    def changeRoadmapStep(self, window, page):
        pageid = self._model.getCurrentPageId()
        if pageid != page:
            if not self._setCurrentPage(window, page):
                self._model.setCurrentPageId(pageid)
        print("WizardManager.changeRoadmapStep() %s" % page)

    def enableButton(self, window, button, enabled):
        self._view.enableButton(window, button, enabled)

    def setDefaultButton(self, window, button):
        self._view.setDefaultButton(window, button)

    def travelNext(self, window):
        page = self._getNextPage()
        if page is not None:
            return self._setCurrentPage(window, page)
        return False

    def travelPrevious(self, window):
        page = self._getPreviousPage()
        if page is not None:
            return self._setCurrentPage(window, page)
        return False

    def enablePage(self, index, enabled):
        self._model.enablePage(index, enabled)

    def updateTravelUI(self, window):
        self._model.updateRoadmap()
        self._updateButton(window)

    def advanceTo(self, window, page):
        if page in self._model.getRoadmapPath():
            return self._setCurrentPage(window, page)
        return False

    def goBackTo(self, window, page):
        if page in self._model.getRoadmapPath():
            return self._setCurrentPage(window, page)
        return False

    def executeWizard(self, dialog):
        if self._currentPath == -1:
            self._initPath(0, False)
        self._initPage(dialog)
        return dialog.execute()

    def _initPath(self, index, final):
        final, paths = self._getActivePath(index, final)
        self._firstPage = min(paths)
        self._lastPage = max(paths)
        self._model.initRoadmap(self._controller, paths, final)

    def _getActivePath(self, index, final):
        if self._multiPaths:
            paths = self._paths[index] if final else self._getFollowingPath(index)
            self._currentPath = index
            final = paths == self._paths[self._currentPath]
        else:
            final, paths = True, self._paths
        return final, paths

    def _getFollowingPath(self, index):
        paths = []
        i = 0
        pageid = self._model.getCurrentPageId()
        for page in self._paths[index]:
            if page > pageid:
                for j in range(len(self._paths)):
                    if j in (index, self._currentPath):
                        continue
                    if i >= len(self._paths[j]) or page != self._paths[j][i]:
                        return tuple(paths)
            paths.append(page)
            i += 1
        return tuple(paths)

    def _initPage(self, window):
        self._setPageStep(window, self._firstPage)
        nextpage = self._isAutoLoad()
        while nextpage and self._canAdvance():
            nextpage = self._initNextPage(window)

    def _isAutoLoad(self, page=None):
        nextindex = 1 if page is None else self.getCurrentPath().index(page) + 1
        return nextindex < self._auto

    def _initNextPage(self, window):
        page = self._getNextPage()
        if page is not None:
            return self._setCurrentPage(window, page) and self._isAutoLoad(page)
        return False

    def _canAdvance(self):
        return self._controller.canAdvance() and self._model.canAdvance()

    def _getPreviousPage(self):
        path = self._model.getRoadmapPath()
        page = self._model.getCurrentPageId()
        if page in path:
            i = path.index(page) - 1
            if i >= 0:
                return path[i]
        return None

    def _getNextPage(self):
        path = self._model.getRoadmapPath()
        page = self._model.getCurrentPageId()
        if page in path:
            i = path.index(page) + 1
            if i < len(path):
                return path[i]
        return None

    def _setCurrentPage(self, window, page):
        if self._deactivatePage(page):
            self._setPageStep(window, page)
            return True
        return False

    def _setPageStep(self, window, pageid):
        self._view.setDialogStep(window, 0)
        # TODO: PageId can be equal to zero but Model.Step must be > 0
        step = self._getDialogStep(pageid)
        if not self._model.hasPage(pageid):
            parent = window.getPeer()
            page = self._controller.createPage(parent, pageid)
            model = page.Window.getModel()
            self._setModelStep(model, step)
            self._model.addPage(pageid, self._setPageWindow(window, model, page))
        self._model.setCurrentPageId(pageid)
        self._model.updateRoadmap()
        self._updateButton(window)
        self._activatePage(pageid)
        self._view.setDialogStep(window, step)

    def _updateButton(self, window):
        self._view.updateButtonPrevious(window, not self._isFirstPage())
        enabled = self._getNextPage() is not None and self._canAdvance()
        self._view.updateButtonNext(window, enabled)
        enabled = self._isLastPage() and self._canAdvance()
        self._view.updateButtonFinish(window, enabled)

    def _deactivatePage(self, new):
        old = self._model.getCurrentPageId()
        reason = self._getCommitReason(old, new)
        if self._model.deactivatePage(old, reason):
            self._controller.onDeactivatePage(old)
            return True
        return False

    def _getCommitReason(self, old=None, new=None):
        old = self._model.getCurrentPageId() if old is None else old
        new = self._model.getCurrentPageId() if new is None else new
        if old < new:
            return FORWARD
        elif old > new:
            return BACKWARD
        else:
            return FINISH

    def _getDialogStep(self, pageid):
        return pageid + 1

    def _setModelStep(self, model, step):
        model.PositionX = self._model.getRoadmapWidth()
        model.PositionY = 0
        model.Step = step
        for control in model.getControlModels():
            control.Step = step

    def _setPageWindow(self, window, model, page):
        if self._resize:
            self._model.setRoadmapSize(model)
            self._view.setDialogSize(window, model)
            self._resize = False
        window.addControl(model.Name, page.Window)
        return page

    def _activatePage(self, page):
        self._controller.onActivatePage(page)
        self._model.activatePage(page)

    def _isFirstPage(self):
        return self._model.getCurrentPageId() == self._firstPage

    def _isLastPage(self):
        return self._isFinal() and self._model.getCurrentPageId() == self._lastPage

    def _isFinal(self):
        return self._model.isFinal()
