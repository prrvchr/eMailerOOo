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

from .griddatamodel import GridDataModel

from com.sun.star.view.SelectionType import MULTI

from unolib import getDialog

from .spoolerhandler import DialogHandler
from .spoolerhandler import GridHandler

from ..configuration import g_extension


class SpoolerView(unohelper.Base):
    def __init__(self, ctx, manager, parent):
        handler = DialogHandler(manager)
        self._dialog = getDialog(ctx, g_extension, 'SpoolerDialog', handler, parent)
        point = uno.createUnoStruct('com.sun.star.awt.Point', 5, 5)
        size = uno.createUnoStruct('com.sun.star.awt.Size', 390, 170)
        rowset = manager.getRowSet()
        grid = self._getGridControl(ctx, rowset, 'GridControl1', point, size, 'SpoolerGrid')
        handler = GridHandler(manager)
        grid.addSelectionListener(handler)
        manager.executeRowSet(rowset)

# SpoolerView setter methods
    def enableButtonRemove(self, enabled):
        self._getButtonRemove().Model.Enabled = enabled

    def dispose(self):
        self._dialog.dispose()

# SpoolerView getter methods
    def execute(self):
        return self._dialog.execute()

    def getParent(self):
        return self._dialog.getPeer()

# SpoolerView private setter methods
    def _setTitle(self, model):
        title = model.resolveString('SendDialog.Title')
        self._dialog.setTitle(title % model.Email)

# SpoolerView private control methods
    def _getGridControl(self, ctx, rowset, name, point, size, tag):
        model = self._getGridModel(ctx, rowset, name, point, size, tag)
        self._dialog.Model.insertByName(name, model)
        return self._dialog.getControl(name)

    def _getGridModel(self, ctx, rowset, name, point, size, tag):
        data = GridDataModel(ctx, rowset)
        service = 'com.sun.star.awt.grid.UnoControlGridModel'
        model = self._dialog.Model.createInstance(service)
        model.Name = name
        model.PositionX = point.X
        model.PositionY = point.Y
        model.Height = size.Height
        model.Width = size.Width
        model.Tag = tag
        model.GridDataModel = data
        model.ColumnModel = data.ColumnModel
        model.SelectionModel = MULTI
        #model.ShowRowHeader = True
        model.BackgroundColor = 16777215
        return model

    def _getButtonRemove(self):
        return self._dialog.getControl('CommandButton2')
