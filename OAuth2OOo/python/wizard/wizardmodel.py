#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from unolib import getStringResource

from .configuration import g_identifier
from .configuration import g_extension

import traceback


class WizardModel(unohelper.Base):
    def __init__(self, ctx):
        self.ctx = ctx
        self._pages = {}
        self._currentPageId = -1
        self._roadmap = None
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)

    def initWizard(self, window, view):
        return self._setRoadmapModel(window.getModel(), view)

    def setRoadmapSize(self, page):
        self._roadmap.Height = page.Height
        self._roadmap.Width = page.PositionX

    def getRoadmapWidth(self):
        return self._roadmap.Width

    def getCurrentPage(self):
        return self._pages.get(self._currentPageId, None)

    def getCurrentPageId(self):
        return self._currentPageId

    def hasPage(self, pageid):
        return pageid in self._pages

    def addPage(self, pageid, page):
        self._pages[pageid] = page

    def setCurrentPageId(self, pageid):
        self._currentPageId = self._roadmap.CurrentItemID = pageid

    def isComplete(self):
        return self._roadmap.Complete

    def activatePage(self, page):
        if page in self._pages:
            self._pages[page].activatePage()

    def enablePage(self, index, enabled):
        self._roadmap.getByIndex(index).Enabled = enabled

    def canAdvance(self):
        return self._canAdvancePage(self._currentPageId)

    def setPageVisible(self, page, enabled):
        if page in self._pages:
            self._pages[page].Window.setVisible(enabled)

    def deactivatePage(self, page, reason):
        if page in self._pages:
            return self._pages[page].commitPage(reason)
        return False

    def doFinish(self, reason):
        return self._pages[self._currentPageId].commitPage(reason)

    def getRoadmapPath(self, enabled=True):
        paths = []
        for i in range(self._roadmap.getCount()):
            item = self._roadmap.getByIndex(i)
            if not enabled:
                paths.append(item.ID)
            elif item.Enabled:
                paths.append(item.ID)
        return tuple(paths)

    def initRoadmap(self, controller, first, paths, complete):
        # TODO: Fixed: Roadmap item will be initialized as follows:
        # TODO: - Previous and current page will be enabled.
        # TODO: - Subsequent pages will be enabled if the previous page can advance otherwise disabled.
        advance = True
        initialized = self._currentPageId != -1
        pageid = self._currentPageId if initialized else first
        for i in range(self._roadmap.getCount() -1, -1, -1):
            self._roadmap.removeByIndex(i)
        i = 0
        for page in paths:
            item = self._roadmap.createInstance()
            item.ID = page
            item.Label = controller.getPageTitle(page)
            if page <= pageid:
                item.Enabled = True
            elif advance and previous in self._pages:
                item.Enabled = advance = self._canAdvancePage(previous)
            else:
                item.Enabled = False
            self._roadmap.insertByIndex(i, item)
            previous = page
            i += 1
        if initialized:
            self._roadmap.CurrentItemID = pageid
        self._roadmap.Complete = complete

    def updateRoadmap(self, first, clear):
        # TODO: Fixed: Roadmap item will be by default:
        # TODO: - Previous pages will be enabled if 'clear' otherwise left unchanged.
        # TODO: - Current page will be enabled, it's mandatory.
        # TODO: - Subsequent pages will be enabled if the previous page can advance otherwise disabled.
        advance = True
        pageid = self._currentPageId if self._currentPageId != -1 else first
        if self._currentPageId == -1:
            print("WizardModel.updateRoadmap() %s ****************************************" % pageid)
        for i in range(self._roadmap.getCount()):
            item = self._roadmap.getByIndex(i)
            if item.ID < pageid:
                if clear:
                    item.Enabled = True
            elif item.ID == pageid:
                item.Enabled = True
            elif advance and previous in self._pages:
                item.Enabled = advance = self._canAdvancePage(previous)
            else:
                item.Enabled = False
            previous = item.ID

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

    def _setRoadmapModel(self, window, view):
        self._roadmap = window.createInstance('com.sun.star.awt.UnoControlRoadmapModel')
        position, size = view.getRoadmapPosition(), view.getRoadmapSize()
        self._roadmap.Name = view.getRoadmapName()
        self._roadmap.PositionX = position.X
        self._roadmap.PositionY = position.Y
        self._roadmap.Height = size.Height
        self._roadmap.Width = size.Width
        self._roadmap.Text = self.resolveString(view.getRoadmapTitle())
        return self._roadmap

    def _canAdvancePage(self, page):
        if page in self._pages:
            return self._pages[page].canAdvance()
        return False
