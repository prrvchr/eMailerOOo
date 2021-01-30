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

import unohelper

from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.frame import XDispatchResultListener

from com.sun.star.frame.DispatchResultState import SUCCESS
from com.sun.star.awt.Key import RETURN

from unolib import executeDispatch

import traceback


class DialogHandler(unohelper.Base,
                    XDialogEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'Send':
            self._manager.sendDocument()
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('Send', )

class WindowHandler(unohelper.Base,
                    XContainerWindowEventHandler):
    def __init__(self, ctx, manager):
        self._ctx = ctx
        self._manager = manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'ChangeSender':
            enabled = event.Source.getSelectedItemPos() != -1
            self._manager.enableRemoveSender(enabled)
            handled = True
        elif method == 'AddSender':
            listener = DispatchListener(self._manager)
            executeDispatch(self._ctx, 'smtp://server', listener)
            handled = True
        elif method == 'RemoveSender':
            self._manager.removeSender()
            handled = True
        elif method == 'EditRecipient':
            control = event.Source
            email = control.getText()
            exist = email in control.getItems()
            self._manager.editRecipient(email, exist)
            handled = True
        elif method == 'ChangeRecipient':
            self._manager.changeRecipient()
            handled = True
        elif method == 'KeyPressed':
            if event.KeyCode == RETURN:
                control = event.Source
                email = control.getText()
                exist = email in control.getItems()
                self._manager.enterRecipient(control, email, exist)
            handled = True
        elif method == 'AddRecipient':
            self._manager.addRecipient()
            handled = True
        elif method == 'RemoveRecipient':
            self._manager.removeRecipient()
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('ChangeSender',
                'AddSender',
                'RemoveSender',
                'EditRecipient',
                'ChangeRecipient',
                'KeyPressed',
                'AddRecipient',
                'RemoveRecipient')


class Page1Handler(unohelper.Base,
                 XContainerWindowEventHandler):
    def __init__(self, manager):
        self._manager = manager

    # XContainerWindowEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'SendAsHtml':
            self._manager.sendAsHtml()
            handled = True
        elif method == 'SendAsAttachment':
            self._manager.sendAsAttachment()
            handled = True
        elif method == 'ViewHtmlDocument':
            self._manager.viewHtmlDocument()
            handled = True
        return handled

    def getSupportedMethodNames(self):
        return ('SendAsHtml',
                'SendAsAttachment',
                'ViewHtmlDocument')


class DispatchListener(unohelper.Base,
                       XDispatchResultListener):
    def __init__(self, manager):
        self._manager = manager

    # XDispatchResultListener
    def dispatchFinished(self, notification):
        if notification.State == SUCCESS:
            self._manager.addSender(notification.Result)

    def disposing(self, source):
        pass
