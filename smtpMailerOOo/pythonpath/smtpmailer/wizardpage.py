#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizardPage
from com.sun.star.util import XRefreshListener
from com.sun.star.sdbc import XRowSetListener
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import PropertySet
from unolib import createService
from unolib import getProperty
from unolib import getStringResource

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
                 XRowSetListener):
    def __init__(self, ctx, id, window, handler):
        try:
            msg = "PageId: %s ..." % id
            self.ctx = ctx
            self.PageId = id
            self.Window = window
            self._handler = handler
            if id == 1:
                control = self.Window.getControl('ListBox1')
                datasources = self._handler.DataSources
                control.Model.StringItemList = datasources
                datasource = self._getDocumentDataSource()
                if datasource in datasources:
                    control.selectItem(datasource, True)
            elif id == 2:
                control = self.Window.getControl('ListBox1')
                control.Model.StringItemList = self._handler.TableNames
                self._handler.addRefreshListener(self)
                self._handler._address.addRowSetListener(self)
                self._handler._recipient.addRowSetListener(self)
                #rowset.addRowSetListener(self)
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
                control = self.Window.getControl('ListBox1')
                control.Model.StringItemList = self._handler.TableNames
            elif self.PageId == 3:
                pass
            print("WizardPage.refreshed(%s) 2" % self.PageId)
        except Exception as e:
            print("WizardPage.refreshed() ERROR: %s - %s" % (e, traceback.print_exc()))
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
        msg = "PageId: %s ..." % self.PageId
        forward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.FORWARD')
        backward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.BACKWARD')
        finish = uno.getConstantByName('com.sun.star.ui.dialogs.WizardTravelType.FINISH')
        self.Window.setVisible(False)
        if self.PageId == 1 and reason == forward:
            pass
        elif self.PageId == 2:
            pass
        elif self.PageId == 3:
            pass
        msg += " Done"
        logMessage(self.ctx, INFO, msg, 'WizardPage', 'commitPage()')
        return True

    def canAdvance(self):
        advance = False
        if self.PageId == 1:
            control = self.Window.getControl('ListBox1')
            advance = control.SelectedItem != ''
        elif self.PageId == 2:
            control = self.Window.getControl('ListBox3')
            advance = control.ItemCount != 0
        elif self.PageId == 3:
            pass
        return advance

    def _getDocumentDataSource(self):
        desktop = 'com.sun.star.frame.Desktop'
        setting = 'com.sun.star.document.Settings'
        document = createService(self.ctx, desktop).CurrentComponent
        if document.supportsService('com.sun.star.text.TextDocument'):
            return document.createInstance(setting).CurrentDatabaseDataSource
        return None

    # XRowSetListener
    def disposing(self, event):
        pass
    def cursorMoved(self, event):
        pass
    def rowChanged(self, event):
        pass
    def rowSetChanged(self, event):
        try:
            print("wizardpage.rowSetChanged() 1")
            if self.PageId == 2:
                if event.Source == self._handler._address:
                    address = self._handler._address
                    self.Window.getControl("ListBox2").Model.StringItemList = getRowResult(address, 0)
                    self.Window.getControl("CommandButton2").Model.Enabled = (address.RowCount != 0)
                    self.Window.getControl("CommandButton3").Model.Enabled = False
                elif event.Source == self._handler._recipient:
                    recipient = self._handler._recipient
                    self.Window.getControl("ListBox3").Model.StringItemList = getRowResult(recipient, 0)
                    self.Window.getControl("CommandButton4").Model.Enabled = False
                    self.Window.getControl("CommandButton5").Model.Enabled = (recipient.RowCount != 0)
            print("wizardpage.rowSetChanged() 2")
        except Exception as e:
            print("wizardpage.rowSetChanged() ERROR: %s - %s" % (e, traceback.print_exc()))

    def _getPropertySetInfo(self):
        properties = {}
        readonly = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.READONLY')
        transient = uno.getConstantByName('com.sun.star.beans.PropertyAttribute.TRANSIENT')
        properties['PageId'] = getProperty('PageId', 'short', readonly)
        properties['Window'] = getProperty('Window', 'com.sun.star.awt.XWindow', readonly)
        return properties
