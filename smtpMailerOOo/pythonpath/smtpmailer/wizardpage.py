#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizardPage
from com.sun.star.util import XRefreshListener
from com.sun.star.container import XContainerListener
from com.sun.star.awt.grid import XGridSelectionListener
from com.sun.star.view.SelectionType import MULTI

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import PropertySet
from unolib import getProperty
from unolib import getStringResource

from .griddatamodel import GridDataModel
from .dbtools import getRowResult
from .logger import logMessage

from .configuration import g_identifier
from .configuration import g_extension
from .configuration import g_column_index

import traceback


class WizardPage(unohelper.Base,
                 PropertySet,
                 XWizardPage,
                 XRefreshListener,
                 XGridSelectionListener):
    def __init__(self, ctx, id, window, handler):
        msg = "PageId: %s loading ..." % id
        print("wizardpage.__init__() 1")
        self.ctx = ctx
        self.PageId = id
        self.Window = window
        self._handler = handler
        print("wizardpage.__init__() 2")
        try:
            if id == 1:
                self._initPage1()
            elif id == 2:
                self._initPage2()
        except Exception as e:
            msg += "Error: %s" % traceback.print_exc()
            logMessage(self.ctx, SEVERE, msg, 'WizardPage', '_initPage%s()' % id)
        print("wizardpage.__init__() 3")
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardPage', '__init__()')

    def _initPage1(self):
        print("wizardpage._initPage1() 1")
        control = self.Window.getControl('ListBox1')
        datasources = self._handler.DataSources
        control.Model.StringItemList = datasources
        datasource = self._handler.getDocumentDataSource()
        print("wizardpage._initPage1() 2 %s-%s" % (datasources, datasource))
        if datasource in datasources:
            print("wizardpage._initPage1() 3")
            self._handler._changeDataSource(self.Window, datasource)
            print("wizardpage._initPage1() 4")
            #self._handler._initColumnsSetting(self.Window)
            print("wizardpage._initPage1() 5")
            self._handler._disabled = True
            print("wizardpage._initPage1() 6")
            control.selectItem(datasource, True)
            print("wizardpage._initPage1() 7")
            self._handler._disabled = False
        print("wizardpage._initPage1() 8")

    def _initPage2(self):
        print("wizardpage._initPage2() 1")
        p = uno.createUnoStruct('com.sun.star.awt.Point', 10, 60)
        s = uno.createUnoStruct('com.sun.star.awt.Size', 115, 115)
        grid1 = self._getGridControl(self._handler._address, 'GridControl1', p, s, 'Addresses')
        grid1.addSelectionListener(self)
        p.X = 160
        grid2 = self._getGridControl(self._handler._recipient, 'GridControl2', p, s, 'Recipients')
        grid2.addSelectionListener(self)
        self._handler.addRefreshListener(self)
        self._handler._recipient.execute()
        self._refreshPage2()
        print("wizardpage._initPage2() 2")

    # XRefreshListener
    def refreshed(self, event):
        tag = event.Source.Model.Tag
        if self.PageId == 1:
            pass
        elif self.PageId == 2 and tag == 'DataSource':
            self._refreshPage2()
        elif self.PageId == 3:
            pass

    # XGridSelectionListener
    def selectionChanged(self, event):
        tag = event.Source.Model.Tag
        enabled = event.Source.hasSelectedRows()
        if tag == 'Addresses':
            self.Window.getControl("CommandButton2").Model.Enabled = enabled
        elif tag == 'Recipients':
            self.Window.getControl("CommandButton3").Model.Enabled = enabled
            if enabled:
                index = event.Source.getSelectedRows()[0]
                if index != self._handler.index:
                    self._handler.setDocumentRecord(index)

    # XContainerListener
    def elementInserted(self, event):
        print("WizardPage.elementInserted()")
        mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
        mri.inspect(event)
    def elementRemoved(self, event):
        print("WizardPage.elementRemoved()")
        mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
        mri.inspect(event)
    def elementReplaced(self, event):
        print("WizardPage.elementReplaced()")
        mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
        mri.inspect(event)

    # XRefreshListener, XGridSelectionListener, XContainerListener
    def disposing(self, event):
        pass

    # XWizardPage
    def activatePage(self):
        # TODO: LibreOffice displays only the first page of the path if you do not manually manage
        # TODO: the visibility of pages on XWizardPage.activatePage() and XWizardPage.commitPage()
        # TODO: reported: Bug 132661 https://bugs.documentfoundation.org/show_bug.cgi?id=132661
        #self.Window.setVisible(True)
        #self._handler._wizard.enablePage(self.PageId +1, True)
        msg = "PageId: %s ..." % self.PageId
        if self.PageId == 1:
            pass
        elif self.PageId == 2:
            pass
        elif self.PageId == 3:
            pass
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardPage', 'activatePage()')

    def commitPage(self, reason):
        try:
            msg = "PageId: %s ..." % self.PageId
            forward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.FORWARD')
            backward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.BACKWARD')
            finish = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.FINISH')
            # TODO: LibreOffice displays only the first page of the path if you do not manually manage
            # TODO: the visibility of pages on XWizardPage.activatePage() and XWizardPage.commitPage()
            # TODO: reported: Bug 132661 https://bugs.documentfoundation.org/show_bug.cgi?id=132661
            #self.Window.setVisible(False)
            #if reason == backward and not self.canAdvance():
            #    self._handler._wizard.enablePage(self.PageId +1, False)
            if self.PageId == 1:
                if self._handler._modified:
                    self._handler.saveSetting(self.Window)
                    self._handler._database.DatabaseDocument.store()
            elif self.PageId == 2:
                print("wizardpage.commitPage() 1")
                #mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
                #mri.inspect(self._handler._database)
                if self._handler._modified:
                    self._handler.saveSelection()
                    self._handler._database.DatabaseDocument.store()
                print("wizardpage.commitPage() 2")
            elif self.PageId == 3:
                pass
            msg += " Done"
            logMessage(self.ctx, INFO, msg, 'WizardPage', 'commitPage()')
            return True
        except Exception as e:
            print("WizardPage.commitPage() ERROR: %s - %s" % (e, traceback.print_exc()))

    def canAdvance(self):
        advance = False
        #print("wizardpage.canAdvance() 1 %s" % advance)
        if self.PageId == 1:
            advance = self._handler.Connection is not None
            advance &= self.Window.getControl("ListBox5").ItemCount != 0
        elif self.PageId == 2:
            advance = self._handler._recipient.RowCount != 0
        elif self.PageId == 3:
            pass
        #print("wizardpage.canAdvance() 2 %s" % advance)
        return advance

    def _getGridControl(self, rowset, name, point, size, tag):
        dialog = self.Window.getModel()
        model = self._getGridModel(rowset, dialog, name, point, size, tag)
        dialog.insertByName(name, model)
        return self.Window.getControl(name)

    def _getGridModel(self, rowset, dialog, name, point, size, tag):
        data = GridDataModel(self.ctx, rowset)
        model = dialog.createInstance('com.sun.star.awt.grid.UnoControlGridModel')
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

    def _refreshPage2(self):
        self._handler.refreshTables(self.Window.getControl('ListBox1'))

    def _getPropertySetInfo(self):
        properties = {}
        readonly = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.READONLY')
        transient = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.TRANSIENT')
        properties['PageId'] = getProperty('PageId', 'short', readonly)
        properties['Window'] = getProperty('Window', 'com.sun.star.awt.XWindow', readonly)
        return properties
