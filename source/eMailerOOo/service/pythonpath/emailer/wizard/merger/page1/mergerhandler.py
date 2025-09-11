#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-25 https://prrvchr.github.io                                  ║
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

import unohelper

from com.sun.star.awt import XContainerWindowEventHandler

from com.sun.star.awt.Key import RETURN

import traceback


class WindowHandler(unohelper.Base,
                    XContainerWindowEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, window, event, method):
        try:
            handled = False
            # TODO: During WizardPage initialization the listener must be disabled...
            enabled = self._manager.isHandlerEnabled()
            if method == 'ChangeAddressBook':
                if enabled:
                    control = event.Source
                    addressbook = control.getSelectedItem()
                    self._manager.changeAddressBook(addressbook)
                handled = True
            elif method == 'NewAddressBook':
                self._manager.newAddressBook()
                handled = True
            elif method == 'ChangeAddressBookTable':
                if enabled:
                    control = event.Source
                    if control.getSelectedItemPos() != -1:
                        table = control.getSelectedItem()
                        self._manager.changeAddressBookTable(table)
                handled = True
            elif method == 'ChangeAddressBookColumn':
                if enabled:
                    control = event.Source
                    if control.getSelectedItemPos() != -1:
                        column = control.getSelectedItem()
                        self._manager.changeAddressBookColumn(column)
                handled = True
            elif method == 'EditQuery':
                control = event.Source
                if control.Model.Enabled:
                    subquery = None
                    query = control.getText().strip()
                    queries = control.getItems()
                    exist = query in queries
                    if exist:
                        index = queries.index(query)
                        subquery = control.Model.getItemData(index)
                    self._manager.editQuery(query, subquery, exist)
                handled = True
            elif method == 'EnterQuery':
                if event.KeyCode == RETURN:
                    control = event.Source
                    query = control.getText().strip()
                    if not query in control.getItems():
                        self._manager.enterQuery(query)
                handled = True
            elif method == 'AddQuery':
                self._manager.addQuery()
                handled = True
            elif method == 'RemoveQuery':
                self._manager.removeQuery()
                handled = True
            elif method == 'ChangeEmail':
                control = event.Source
                imax = control.ItemCount -1
                position = control.getSelectedItemPos()
                self._manager.changeEmail(imax, position)
                handled = True
            elif method == 'AddEmail':
                self._manager.addEmail()
                handled = True
            elif method == 'UpEmail':
                self._manager.upEmail()
                handled = True
            elif method == 'DownEmail':
                self._manager.downEmail()
                handled = True
            elif method == 'RemoveEmail':
                self._manager.removeEmail()
                handled = True
            elif method == 'ChangeIdentifier':
                control = event.Source
                imax = control.ItemCount -1
                position = control.getSelectedItemPos()
                self._manager.changeIdentifier(imax, position)
                handled = True
            elif method == 'AddIdentifier':
                self._manager.addIdentifier()
                handled = True
            elif method == 'UpIdentifier':
                self._manager.upIdentifier()
                handled = True
            elif method == 'DownIdentifier':
                self._manager.downIdentifier()
                handled = True
            elif method == 'RemoveIdentifier':
                self._manager.removeIdentifier()
                handled = True
            return handled
        except Exception as e:
            msg = "Error: %s" % traceback.format_exc()
            print(msg)

    def getSupportedMethodNames(self):
        return ('ChangeAddressBook',
                'NewAddressBook',
                'ChangeAddressBookTable',
                'ChangeAddressBookColumn',
                'EditQuery',
                'EnterQuery',
                'AddQuery',
                'RemoveQuery',
                'ChangeEmail',
                'AddEmail',
                'UpEmail',
                'DownEmail',
                'RemoveEmail',
                'ChangeIdentifier',
                'AddIdentifier',
                'UpIdentifier',
                'DownIdentifier',
                'RemoveIdentifier')
