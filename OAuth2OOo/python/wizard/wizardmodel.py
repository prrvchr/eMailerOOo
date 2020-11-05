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
        self._disabledPages = []
        self._roadmap = None
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)

    def initWizard(self, view):
        return self._setRoadmapModel(view)

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

    def enablePage(self, pageid, enabled):
        if enabled:
            if pageid in self._disabledPages:
                self._disabledPages.remove(pageid)
        elif pageid not in self._disabledPages:
            self._disabledPages.append(pageid)

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

    def getActivePath(self, path):
        return tuple(page for page in path if page not in self._disabledPages)

    def getRoadmapPath(self):
        paths = []
        for i in range(self._roadmap.getCount()):
            item = self._roadmap.getByIndex(i)
            if item.Enabled:
                paths.append(item.ID)
        return tuple(paths)

    def initRoadmap(self, controller, first, paths, complete):
        # TODO: Fixed: Roadmap item will be initialized as follows:
        # TODO: - Previous page will be, if not explicitly disabled, enabled.
        # TODO: - Current page will be enabled, it's mandatory.
        # TODO: - Subsequent pages will be, if not explicitly disabled, enabled
        # TODO:   if the previous page can advance, otherwise disabled.
        advance = True
        previous = -1
        initialized = self._currentPageId != -1
        pageid = self._currentPageId if initialized else first
        for i in range(self._roadmap.getCount() -1, -1, -1):
            self._roadmap.removeByIndex(i)
        i = 0
        for page in paths:
            enabled = self._isPageEnabled(page)
            item = self._roadmap.createInstance()
            item.ID = page
            item.Label = controller.getPageTitle(page)
            if page < pageid:
                item.Enabled = enabled
            elif page == pageid:
                item.Enabled = True
            elif not enabled:
                item.Enabled = False
            elif advance and previous in self._pages:
                advance = self._canAdvancePage(previous)
                item.Enabled = advance
            else:
                item.Enabled = False
            self._roadmap.insertByIndex(i, item)
            if enabled:
                previous = page
            i += 1
        if initialized:
            self._roadmap.CurrentItemID = pageid
        self._roadmap.Complete = complete

    def updateRoadmap(self, first):
        # TODO: Fixed: Roadmap item will be by default:
        # TODO: - Previous pages will be, if not explicitly disabled, enabled.
        # TODO: - Current page will be enabled, it's mandatory.
        # TODO: - Subsequent pages will be, if not explicitly disabled, enabled 
        # TODO:   if the previous page can advance otherwise disabled.
        advance = True
        previous = -1
        pageid = self._currentPageId if self._currentPageId != -1 else first
        if self._currentPageId == -1:
            print("WizardModel.updateRoadmap() %s ****************************************" % pageid)
        for i in range(self._roadmap.getCount()):
            item = self._roadmap.getByIndex(i)
            page = item.ID
            enabled = self._isPageEnabled(page)
            if page < pageid:
                item.Enabled = enabled
            elif page == pageid:
                item.Enabled = True
            elif not enabled:
                item.Enabled = False
            elif advance and previous in self._pages:
                advance = self._canAdvancePage(previous)
                item.Enabled = advance
            else:
                item.Enabled = False
            if enabled:
                previous = page

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

    def _isPageEnabled(self, page):
        return page not in self._disabledPages

    def _setRoadmapModel(self, view):
        window = view.DialogWindow.Model
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
