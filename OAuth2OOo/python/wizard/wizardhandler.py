#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt import XItemListener

import traceback


class DialogHandler(unohelper.Base,
                    XDialogEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'Help':
            handled = True
        elif method == 'Previous':
            self._manager.travelPrevious()
            handled = True
        elif method == 'Next':
            self._manager.travelNext()
            handled = True
        elif method == 'Finish':
            self._manager.doFinish(dialog)
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('Help', 'Previous', 'Next', 'Finish')


class ItemHandler(unohelper.Base,
                  XItemListener):
    def __init__(self, manager):
        self._manager = manager

    # XItemListener
    def itemStateChanged(self, event):
        self._manager.changeRoadmapStep(event.ItemId)

    def disposing(self, event):
        pass
