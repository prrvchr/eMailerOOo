#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt import XItemListener

import traceback


class DialogHandler(unohelper.Base,
                    XDialogEventHandler):
    def __init__(self, ctx, manager):
        self.ctx = ctx
        self._manager = manager
        print("DialogHandler.__init__()")

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        print("DialogHandler.callHandlerMethod() %s" % method)
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


class ItemHandler(unohelper.Base,
                  XItemListener):
    def __init__(self, ctx, dialog, manager):
        self.ctx = ctx
        self._dialog = dialog
        self._manager = manager
        print("ItemHandler.__init__()")

    # XItemListener
    def itemStateChanged(self, event):
        page = event.ItemId
        self._manager.changeRoadmapStep(self._dialog, page)

    def disposing(self, event):
        pass
