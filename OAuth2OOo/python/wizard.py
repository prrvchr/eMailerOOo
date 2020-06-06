#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizard
from com.sun.star.lang import XServiceInfo
from com.sun.star.lang import XInitialization
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt import XItemListener

from com.sun.star.lang import IllegalArgumentException
from com.sun.star.util import InvalidStateException
from com.sun.star.container import NoSuchElementException

from com.sun.star.ui.dialogs.WizardTravelType import FORWARD
from com.sun.star.ui.dialogs.WizardTravelType import BACKWARD
from com.sun.star.ui.dialogs.WizardTravelType import FINISH

from com.sun.star.ui.dialogs.WizardButton import NEXT
from com.sun.star.ui.dialogs.WizardButton import PREVIOUS
from com.sun.star.ui.dialogs.WizardButton import FINISH
from com.sun.star.ui.dialogs.WizardButton import CANCEL
from com.sun.star.ui.dialogs.WizardButton import HELP

from unolib import getDialog
from unolib import createService
from unolib import getStringResource

from .configuration import g_identifier
from .configuration import g_extension

from .logger import getMessage

import traceback


class Wizard(unohelper.Base,
             XWizard,
             XInitialization,
             XItemListener,
             XDialogEventHandler):
    def __init__(self, ctx):
        self.ctx = ctx
        self._pages = {}
        self._paths = ()
        self._currentPath = -1
        self._firstPage = -1
        self._lastPage = -1
        self._multiPath = False
        self._controller = None
        self._helpUrl = ''
        self._currentItemID = -1
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)
        self._dialog = getDialog(self.ctx, g_extension, 'Wizard', self)
        point = uno.createUnoStruct('com.sun.star.awt.Point', 0, 0)
        size = uno.createUnoStruct('com.sun.star.awt.Size', 85, 180)
        roadmap = self._getRoadmapControl('RoadmapControl1', point, size)
        roadmap.addItemListener(self)

    @property
    def _roadMap(self):
        return self._dialog.getControl('RoadmapControl1').getModel()
    @property
    def _currentPage(self):
        return self._roadMap.CurrentItemID
    @_currentPage.setter
    def _currentPage(self, page):
        self._currentItemID = page
        self._roadMap.CurrentItemID = page
    @property
    def _isFinal(self):
        return self._roadMap.Complete

    @property
    def HelpURL(self):
        return self._helpUrl
    @HelpURL.setter
    def HelpURL(self, url):
        self._helpUrl = url
        self._dialog.getControl('CommandButton1').Model.Enabled = url != ''
    @property
    def DialogWindow(self):
        return self._dialog

    # XInitialization
    def initialize(self, args):
        code = self._getInvalidArgumentCode(args)
        if code != -1:
            raise self._getIllegalArgumentException(0, code)
        paths = args[0]
        code = self._getInvalidPathsCode(paths)
        if code != -1:
            raise self._getIllegalArgumentException(0, code)
        self._multiPath = isinstance(paths[0], tuple)
        self._paths = paths
        controller = args[1]
        code = self._getInvalidControllerCode(controller)
        if code != -1:
            raise self._getIllegalArgumentException(0, code)
        self._controller = controller

    # XItemListener
    def itemStateChanged(self, event):
        old = self._currentItemID
        new = event.ItemId
        if self._setPage(old, new):
            self.updateTravelUI()
        else:
            self._currentPage = old
    def disposing(self, event):
        pass

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'Help':
            handled = True
        elif method == 'Next':
            self.travelNext()
            handled = True
        elif method == 'Previous':
            self.travelPrevious()
            handled = True
        elif method == 'Finish':
            handled = self._doFinish()
        return handled
    def getSupportedMethodNames(self):
        return ('Help', 'Previous', 'Next', 'Finish')

    # XWizard
    def getCurrentPage(self):
        page = self._currentPage
        if page in self._pages:
            return self._pages[page]
        return None

    def enableButton(self, button, enabled):
        if button == HELP:
            self._dialog.getControl('CommandButton1').Model.Enabled = enabled
        elif button == PREVIOUS:
            self._dialog.getControl('CommandButton2').Model.Enabled = enabled
        elif button == NEXT:
            self._dialog.getControl('CommandButton3').Model.Enabled = enabled
        elif button == FINISH:
            self._dialog.getControl('CommandButton4').Model.Enabled = enabled
        elif button == CANCEL:
            self._dialog.getControl('CommandButton5').Model.Enabled = enabled

    def setDefaultButton(self, button):
        if button == HELP:
            self._dialog.getControl('CommandButton1').Model.DefaultButton = True
        elif button == PREVIOUS:
            self._dialog.getControl('CommandButton2').Model.DefaultButton = True
        elif button == NEXT:
            self._dialog.getControl('CommandButton3').Model.DefaultButton = True
        elif button == FINISH:
            self._dialog.getControl('CommandButton4').Model.DefaultButton = True
        elif button == CANCEL:
            self._dialog.getControl('CommandButton5').Model.DefaultButton = True

    def travelNext(self):
        print("Wizard.travelNext()")
        old, new = self._getNextPage()
        if old is not None:
            return self._setCurrentPage(old, new)
        return False

    def travelPrevious(self):
        print("Wizard.travelPrevious()")
        old, new = self._getPreviousPage()
        if old is not None:
            return self._setCurrentPage(old, new)
        return False

    def enablePage(self, page, enabled):
        if page == self._currentPage:
            raise self._getInvalidStateException(102)
        path = self._getPath(False)
        if page not in path:
            raise self._getNoSuchElementException(103)
        index = path.index(page)
        self._roadMap.getByIndex(index).Enabled = enabled

    def updateTravelUI(self):
        self._updateRoadmapItem()
        self._updateControlButton()

    def advanceTo(self, page):
        print("Wizard.advanceTo()")
        if page in self._getPath():
            return self._setCurrentPage(self._currentPage, page)
        return False

    def goBackTo(self, page):
        print("Wizard.goBackTo()")
        if page in self.self._getPath():
            return self._setCurrentPage(self._currentPage, page)
        return False

    def activatePath(self, index, final):
        if not self._multiPath:
            return
        if index not in range(len(self._paths)):
            raise self._getNoSuchElementException(104)
        if self._currentPath != -1:
            old = self._paths[self._currentPath]
            new = self._paths[index]
            commun = self._getCommunPaths((old, new))
            if self._currentPage not in commun:
                raise self._getInvalidStateException(105)
        print("Wizard.activatePath()")

    # XExecutableDialog
    def setTitle(self, title):
        self._dialog.setTitle(title)
    def execute(self):
        if self._currentPage == -1:
            print("Wizard.execute() 1")
            self._initPaths(0)
            #self._updateRoadmapItem()
            print("Wizard.execute() 2")
        return self._dialog.execute()

    # XServiceInfo
    def supportsService(self, service):
        pass
    def getImplementationName(self):
        pass
    def getSupportedServiceNames(self):
        pass

    def _getPreviousPage(self):
        page = self._currentPage
        path = self._getPath()
        if page in path:
            i = path.index(page) - 1
            if i >= 0:
                return page, path[i]
        return None, None

    def _getNextPage(self):
        page = self._currentPage
        path = self._getPath()
        if page in path:
            i = path.index(page) + 1
            if i < len(path):
                return page, path[i]
        return None, None

    def _doFinish(self):
        page = self._currentPage
        if self._isFinal and page == self._lastPage:
            reason = self._getCommitReason(page, page)
            if self._pages[page].commitPage(reason):
                if self._controller.confirmFinish():
                    self._dialog.endExecute()
        return True

    def _setCurrentPage(self, old, new):
        if self._setPage(old, new):
            self.updateTravelUI()
            return True
        return False

    def _initPaths(self, index):
        try:
            if self._multiPath:
                final, path = self._initMultiplePathsWizard(index)
            else:
                final, path = self._initSinglePathWizard()
            page = min(path)
            self._firstPage = page
            self._lastPage = max(path)
            self._initRoadmapItem(path, final)
            self._initPage(page)
            self.updateTravelUI()
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def _initSinglePathWizard(self):
        self._currentPath = -1
        return True, self._paths

    def _initMultiplePathsWizard(self, index):
        path = self._getActivePath(self._paths)
        print("Wizard._initMultiPaths() %s" % (path, ))
        self._currentPath = index
        final = path == self._paths[self._currentPath]
        return final, path

    def _getActivePath(self, paths, index=-1):
        if index != -1:
            return self._getUniquePaths(paths, index)
        else:
            return self._getCommunPaths(paths)

    def _getUniquePaths(self, paths, index):
        commun = []
        i = 0
        for p in paths[index]:
            for j in range(0, len(paths)):
                if p != paths[j][i]:
                    return tuple(commun)
            commun.append(p)
            i += 1
        return tuple(commun)

    def _getCommunPaths(self, paths):
        commun = []
        i = 0
        for p in paths[0]:
            for j in range(1, len(paths)):
                if p != paths[j][i]:
                    return tuple(commun)
            commun.append(p)
            i += 1
        return tuple(commun)

    def _getPath(self, enabled=True):
        path = []
        roadmap = self._roadMap
        for i in range(roadmap.getCount()):
            item = roadmap.getByIndex(i)
            if not enabled:
                path.append(item.ID)
            elif item.Enabled:
                path.append(item.ID)
        return tuple(path)

    def _initRoadmapItem(self, path, final):
        roadmap = self._roadMap
        i = 0
        initialized = roadmap.CurrentItemID != -1
        for page in path:
            item = roadmap.createInstance()
            item.ID = page
            item.Label = self._controller.getPageTitle(page)
            if i != 0:
                item.Enabled = initialized and self._canPageAdvance(id)
            roadmap.insertByIndex(i, item)
            id = page
            i += 1
        roadmap.Complete = final

    def _updateRoadmapItem(self):
        roadmap = self._roadMap
        for i in range(roadmap.getCount()):
            item = roadmap.getByIndex(i)
            if i == 0:
                item.Enabled = True
            elif id in self._pages:
                item.Enabled = self._canPageAdvance(id)
            else:
                item.Enabled = False
            id = item.ID
        
    def _updateControlButton(self):
        page = self._currentPage
        enabled = self._getPreviousPage() != (None, None)
        self._dialog.getControl('CommandButton2').Model.Enabled = enabled
        enabled = self._getNextPage()!= (None, None) and self._canAdvance(page)
        self._dialog.getControl('CommandButton3').Model.Enabled = enabled
        enabled = self._isFinal and page == self._lastPage
        self._dialog.getControl('CommandButton4').Model.Enabled = enabled

    def _canAdvance(self, page):
        advance = self._controller.canAdvance()
        if page in self._pages:
            advance &= self._pages[page].canAdvance()
        return advance

    def _canPageAdvance(self, page):
        advance = False
        if page in self._pages:
            advance = self._pages[page].canAdvance()
        return advance

    def _setPage(self, old, new):
        if self._deactivatePage(old, new):
            self._initPage(new)
            return True
        return False

    def _initPage(self, id):
        self._setDialogStep(0)
        # TODO: PageId can be equal to zero but Model.Step must be > 0
        step = self._getDialogStep(id)
        if id not in self._pages:
            parent = self._dialog.getPeer()
            page = self._controller.createPage(parent, id)
            model = page.Window.getModel()
            self._setModelStep(model, step)
            self._dialog.addControl(model.Name, page.Window)
            #self._dialog.getModel().insertByName(model.Name, model)
            self._pages[id] = page
        self._currentPage = id
        self._setDialogStep(step)
        self._activatePage(id)

    def _setModelStep(self, model, step):
        model.PositionX = self._roadMap.Width
        model.PositionY = 0
        model.Step = step
        for control in model.getControlModels():
            control.Step = step

    def _deactivatePage(self, old, new):
        if old in self._pages:
            if self._pages[old].commitPage(self._getCommitReason(old, new)):
                self._controller.onDeactivatePage(old)
                return True
        return False

    def _getCommitReason(self, old, new):
        if old < new:
            return FORWARD
        elif old > new:
            return BACKWARD
        else:
            return FINISH

    def _activatePage(self, page):
        if page in self._pages:
            self._controller.onActivatePage(page)
            self._pages[page].activatePage()

    def _getDialogStep(self, id):
        return id + 1

    def _setDialogStep(self, step):
        self._dialog.getModel().Step = step

    def _getRoadmapControl(self, name, point, size, tag=None):
        dialog = self._dialog.getModel()
        model = self._getRoadmapModel(dialog, name, point, size, tag)
        dialog.insertByName(name, model)
        return self._dialog.getControl(name)

    def _getRoadmapModel(self, dialog, name, point, size, tag):
        model = dialog.createInstance('com.sun.star.awt.UnoControlRoadmapModel')
        model.Name = name
        model.PositionX = point.X
        model.PositionY = point.Y
        model.Height = size.Height
        model.Width = size.Width
        if tag is not None:
            model.Tag = tag
        model.Text = self._stringResource.resolveString('Wizard.Roadmap.Text')
        model.Border = 0
        return model

    def _getInvalidArgumentCode(self, args):
        code = -1
        if not isinstance(args, tuple) or len(args) != 2:
            code = 100
        return code

    def _getInvalidPathsCode(self, paths):
        code = -1
        if not isinstance(paths, tuple) or len(paths) < 1:
             code = 101
        return code

    def _getInvalidControllerCode(self, controller):
        return -1

    def _getIllegalArgumentException(self, position, code):
        e = IllegalArgumentException()
        e.ArgumentPosition = position
        e.Message = getMessage(self.ctx, code)
        e.Context = self
        return e

    def _getInvalidStateException(self, code):
        e = InvalidStateException()
        e.Message = getMessage(self.ctx, code)
        e.Context = self
        return e

    def _getNoSuchElementException(self, code):
        e = NoSuchElementException()
        e.Message = getMessage(self.ctx, code)
        e.Context = self
        return e
