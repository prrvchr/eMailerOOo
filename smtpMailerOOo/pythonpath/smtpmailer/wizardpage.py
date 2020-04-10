#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizardPage
from com.sun.star.util import XRefreshListener
from com.sun.star.container import XContainerListener
from com.sun.star.awt.grid import XGridSelectionListener
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import PropertySet
from unolib import createService
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
        try:
            msg = "PageId: %s ..." % id
            self.ctx = ctx
            self.PageId = id
            self.Window = window
            self._handler = handler
            if id == 1:
                #dbcontext = createService(self.ctx, 'com.sun.star.sdb.DatabaseContext')
                #dbcontext.addContainerListener(self)
                control = self.Window.getControl('ListBox1')
                datasources = self._handler.DataSources
                control.Model.StringItemList = datasources
                datasource = self._handler.getDocumentDataSource()
                if datasource in datasources:
                    control.selectItem(datasource, True)
            elif id == 2:
                print("wizardpage.__init__() 1")
                self._refreshPage2()
                self._handler.addRefreshListener(self)
                self._initGrid(self.Window.getControl('GridControl1'), self._handler._address)
                self._initGrid(self.Window.getControl('GridControl2'), self._handler._recipient)

                #mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
                #mri.inspect(grid1)
                print("wizardpage.__init__() 2")
            elif id == 3:
                pass
            msg += " Done"
            logMessage(self.ctx, INFO, msg, 'WizardPage', '__init__()')
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            logMessage(self.ctx, SEVERE, msg, 'WizardPage', '__init__()')

    # XRefreshListener
    def refreshed(self, event):
        try:
            tag = event.Source.Model.Tag
            print("WizardPage.refreshed(%s) 1" % self.PageId)
            if self.PageId == 1:
                pass
            elif self.PageId == 2 and tag == 'DataSource':
                #mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
                #mri.inspect(event)
                self._refreshPage2()
            elif self.PageId == 3:
                pass
            print("WizardPage.refreshed(%s) 2" % self.PageId)
        except Exception as e:
            print("WizardPage.refreshed() ERROR: %s - %s" % (e, traceback.print_exc()))

    # XGridSelectionListener
    def selectionChanged(self, event):
        try:
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
        except Exception as e:
            print("WizardPage.selectionChanged() ERROR: %s - %s" % (e, traceback.print_exc()))

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
        self.Window.setVisible(False)
        msg = "PageId: %s ..." % self.PageId
        if self.PageId == 1:
            pass
        elif self.PageId == 2:
            pass
        elif self.PageId == 3:
            pass
        self.Window.setVisible(True)
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardPage', 'activatePage()')

    def commitPage(self, reason):
        try:
            msg = "PageId: %s ..." % self.PageId
            forward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.FORWARD')
            backward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.BACKWARD')
            finish = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.FINISH')
            self.Window.setVisible(False)
            if self.PageId == 1 and reason == forward:
                pass
            elif self.PageId == 2:
                print("wizardpage.commitPage() 1")
                #mri = self.ctx.ServiceManager.createInstance('mytools.Mri')
                #mri.inspect(self._handler._database)
                self._handler._query.Order = self._handler._recipient.Order
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
        if self.PageId == 1:
            control = self.Window.getControl('ListBox1')
            advance = control.SelectedItem != ''
        elif self.PageId == 2:
            control = self.Window.getControl('GridControl2')
            advance = control.Model.GridDataModel.RowCount != 0
        elif self.PageId == 3:
            pass
        return advance

    def _initGrid(self, grid, rowset):
        # TODO: Because we need a GridDataListener who listen change on the GridDataModel
        # TODO: We need to re assign the Model, and not only just set the GridDataModel
        grid.setModel(self._getGridModel(grid.getModel(), rowset))
        grid.addSelectionListener(self)

    def _getGridModel(self, grid, rowset):
        model = GridDataModel(self.ctx, rowset)
        #self.Window.getControl('ListBox2').addActionListener(model)
        #self.Window.getControl('ListBox2').addItemListener(model)
        grid.GridDataModel = model
        grid.ColumnModel = model.ColumnModel
        return grid

    def _refreshPage2(self):
        print("wizardpage._refreshPage2() 1")
        listbox1 = self.Window.getControl('ListBox1')
        tables = self._handler.TableNames
        listbox1.Model.StringItemList = ()
        listbox1.Model.StringItemList = tables
        table = self._handler._query.UpdateTableName
        table = tables[0] if len(tables) != 0 and table == '' else table
        print("wizardpage._refreshPage2() 2")
        listbox1.selectItem(table, True)
        print("wizardpage._refreshPage2() 3")
        self._handler._disabled = True
        listbox2 = self.Window.getControl('ListBox2')
        listbox2.Model.StringItemList = ()
        listbox2.Model.StringItemList = self._handler.ColumnNames
        columns = self._handler.getOrderIndex()
        print("wizardpage._refreshPage2() 4 %s - %s" % (listbox2.getSelectedItemsPos(), columns))
        listbox2.selectItemsPos(columns, True)
        self._handler._disabled = False
        #self._handler._recipient.execute()
        print("wizardpage._refreshPage2() 5")

    def _getPropertySetInfo(self):
        properties = {}
        readonly = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.READONLY')
        transient = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.TRANSIENT')
        properties['PageId'] = getProperty('PageId', 'short', readonly)
        properties['Window'] = getProperty('Window', 'com.sun.star.awt.XWindow', readonly)
        return properties
