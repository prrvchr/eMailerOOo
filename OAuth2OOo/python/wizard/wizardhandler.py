#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt import XItemListener

from unolib import createService

import traceback


class WizardHandler(unohelper.Base,
                    XItemListener,
                    XDialogEventHandler):
    def __init__(self, ctx, manager):
        self.ctx = ctx
        self._manager = manager

    # XItemListener
    def itemStateChanged(self, event):
        page = event.ItemId
        mri = createService(self.ctx, 'mytools.Mri')
        mri.inspect(event.Source)
        self._manager.changeRoadmapStep(page)
        #if self._currentPage != page:
        #    if not self._setPage(self._currentPage, page):
        #        self._getRoadmap().CurrentItemID = self._currentPage

    def disposing(self, event):
        pass

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'Help':
            handled = True
        elif method == 'Previous':
            self._manager.travelPrevious(dialog)
            handled = True
        elif method == 'Next':
            self._manager.travelNext(dialog)
            handled = True
        elif method == 'Finish':
            self._manager.doFinish(dialog)
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('Help', 'Previous', 'Next', 'Finish')
