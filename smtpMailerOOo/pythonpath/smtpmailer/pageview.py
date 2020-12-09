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

from .logger import logMessage

import traceback


class PageView(unohelper.Base):
    def __init__(self, ctx, window):
        self.ctx = ctx
        self.Window = window
        print("PageView.__init__()")

# PageView setter methods
    def initPage1(self, model):
         return self._setDataSources(model)

    def updateControl(self, model, control, tag):
        try:
            if tag == 'DataSource':
                control.Model.StringItemList = model.DataSources
            elif tag == 'Columns':
                button = self._getEmailAddressAdd()
                button.Model.Enabled = self._canAddItem(control, self._getEmailAddress())
                button = self._getPrimaryKeyAdd()
                button.Model.Enabled = self._canAddItem(control, self._getPrimaryKey())
            elif tag == 'EmailAddress':
                imax = control.ItemCount -1
                position = control.getSelectedItemPos()
                self._getRemoveEmailAddress().Model.Enabled = position != -1
                self._getMoveBefore().Model.Enabled = position > 0
                self._getMoveAfter().Model.Enabled = -1 < position < imax
            elif tag == 'PrimaryKey':
                position = control.getSelectedItemPos()
                self._getRemovePrimaryKey().Model.Enabled = position != -1
            elif tag == 'Addresses':
                selected = control.hasSelectedRows()
                enabled = control.Model.GridDataModel.RowCount != 0
                self._getAddAllRecipient().Model.Enabled = enabled
                self._getAddRecipient().Model.Enabled = selected
            elif tag == 'Recipients':
                selected = control.hasSelectedRows()
                enabled = control.Model.GridDataModel.RowCount != 0
                print("WizardHandler._updateControl() %s - %s" % (selected, enabled))
                self._getRemoveRecipient().Model.Enabled = selected
                self._getRemoveAllRecipient().Model.Enabled = enabled
        except Exception as e:
            print("WizardHandler._updateControl() ERROR: %s - %s" % (e, traceback.print_exc()))

# PageView getter methods
    def getPageTitle(self, model, pageid):
        return model.resolveString(self._getPageTitle(pageid))

# PageView private methods
    def _getControlByTag(self, tag):
        if tag == 'Columns':
            control = self._getColumns()
        elif tag == 'EmailAddress':
            control = self._getEmailAddress()
        elif tag == 'PrimaryKey':
            control = self._getPrimaryKey()
        else:
            print("PageView._getControlByTag() Error: '%s' don't exist!!! **************" % tag)
        return control

    def _canAddItem(self, control, listbox):
        enabled = control.getSelectedItemPos() != -1
        if enabled and listbox.ItemCount != 0:
            column = control.getSelectedItem()
            columns = getStringItemList(listbox)
            return column not in columns
        return enabled

# PageView private message methods
    def _getPageTitle(self, pageid):
        return 'PageWizard%s.Title' % pageid

# PageView private getter control methods for Page1
    def _getDataSource(self):
        return self.Window.getControl('ListBox1')

    def _getColumns(self):
        return self.Window.getControl('ListBox2')

    def _getTables(self):
        return self.Window.getControl('ListBox3')

    def _getEmailAddress(self):
        return self.Window.getControl('ListBox4')

    def _getPrimaryKey(self):
        return self.Window.getControl('ListBox5')

    def _getAddEmailAddress(self):
        return self.Window.getControl('CommandButton2')

    def _getRemoveEmailAddress(self):
        return self.Window.getControl('CommandButton3')

    def _getMoveBefore(self):
        return self.Window.getControl('CommandButton4')

    def _getMoveAfter(self):
        return self.Window.getControl('CommandButton5')

    def _getAddPrimaryKey(self):
        return self.Window.getControl('CommandButton6')

    def _getRemovePrimaryKey(self):
        return self.Window.getControl('CommandButton7')


# PageView private getter control methods for Page2
    def _getAddAllRecipient(self):
        return self.Window.getControl('CommandButton1')

    def _getAddRecipient(self):
        return self.Window.getControl('CommandButton2')

    def _getRemoveRecipient(self):
        return self.Window.getControl('CommandButton3')

    def _getRemoveAllRecipient(self):
        return self.Window.getControl('CommandButton4')


# PageView private setter control methods
    def _setDataSources(self, model):
        initialized = False
        datasources = model.DataSources
        control = self._getDataSource()
        control.Model.StringItemList = datasources
        datasource = model.getDocumentDataSource()
        if datasource in datasources:
            if model.setDataSource(datasource):
                document, form = model.getForm(False)
                self._setTables(model, document, 'PrimaryTable')
                self._setEmailAddress(model, document, 'EmailColumns')
                self._setPrimaryKey(model, document, 'IndexColumns')
                if form is not None:
                    form.close()
                self._updateControlByTag(model, 'Columns')
                self._updateControlByTag(model, 'EmailAddress')
                self._updateControlByTag(model, 'PrimaryKey')
                model.refreshControl(self._getDataSource())
                control.selectItem(datasource, True)
                initialized = True
        return initialized

    def _setTables(self, model, document, property):
        control = self._getTables()
        control.Model.StringItemList = model.TableNames
        table = model.getDocumentProperty(document, property)
        if table is None:
            control.selectItemPos(0, True)
        else:
            control.selectItem(table, True)

    def _setEmailAddress(self, model, document, property):
        self._getEmailAddress().Model.StringItemList = model.getEmailColumn(document, property)

    def _setPrimaryKey(self, model, document, property):
        self._getPrimaryKey().Model.StringItemList = model.getIndexColumns(document, property)

    def _updateControlByTag(self, model, tag):
        control = self._getControlByTag(tag)
        self.updateControl(model, control, tag)
