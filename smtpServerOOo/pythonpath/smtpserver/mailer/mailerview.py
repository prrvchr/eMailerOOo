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

from unolib import getDialog
from unolib import getContainerWindow

from .mailerhandler import DialogHandler
from .mailerhandler import WindowHandler
from .mailerhandler import Page1Handler
from .mailerhandler import Page2Handler

from smtpserver import g_extension

import traceback


class MailerView(unohelper.Base):
    def __init__(self, ctx, parent, manager):
        handler = DialogHandler(manager)
        self._dialog = getDialog(ctx, g_extension, 'MailerDialog', handler, parent)
        parent = self._dialog.getPeer()
        handler = WindowHandler(ctx, manager)
        self._window = getContainerWindow(ctx, parent, handler, g_extension, 'MailerWindow')
        self._window.setVisible(True)
        point = uno.createUnoStruct('com.sun.star.awt.Point', 0, 100)
        size = uno.createUnoStruct('com.sun.star.awt.Size', 285, 85)
        title1 = self._getTabPageTitle(manager.Model, 1)
        title2 = self._getTabPageTitle(manager.Model, 2)
        page1, page2 = self._getTabPages('Tab1', point, size, title1, title2, 1)
        parent = page1.getPeer()
        handler = Page1Handler(manager)
        self._page1 = getContainerWindow(ctx, parent, handler, g_extension, 'MailerPage1')
        self._page1.setVisible(True)
        parent = page2.getPeer()
        handler = Page2Handler(manager)
        self._page2 = getContainerWindow(ctx, parent, handler, g_extension, 'MailerPage2')
        self._page2.setVisible(True)

# MailerView setter methods
    def setRecipient(self, recipients):
        if len(recipients) > 0:
            control = self._getRecipientsList()
            control.Model.StringItemList = recipients
            control.setText(recipients[0])

    def setSenders(self, senders):
        if len(senders) > 0:
            control = self._getSendersList()
            control.Model.StringItemList = senders
            control.selectItemPos(0, True)

    def setTitle(self, title):
        self._dialog.setTitle(title)

    def setSaveSubject(self, state):
        self._getSaveSubject().Model.State = state

    def setSubject(self, subject):
        self._getSubject().Text = subject

    def setSendMode(self, state):
        if state:
            self._getSendAsHtml().Model.State = 1
        else:
            self._getSendAsAttachment().Model.State = 1
            self.setStep(2)

    def setDocumentLabel(self, label):
        self._getDocumentLabel().setText(label)

    def setSaveMessage(self, state):
        self._getSaveMessage().Model.State = state

    def setMessage(self, message):
        self._getMessage().Text = message

    def setAttachMode(self, state):
        try:
            self._getAttachMode().Model.State = state
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def setSaveAttachments(self, state):
        self._getSaveAttachments().Model.State = state

    def setAttachments(self, attachments):
        self._getAttachments().Model.StringItemList = attachments

    def enableButtonViewHtml(self):
        self._getButtonViewHtml().Model.Enabled = True

    def addSender(self, sender):
        control = self._getSendersList()
        count = control.getItemCount()
        if sender not in control.getItems():
            control.addItem(sender, count)
            control.selectItemPos(count, True)

    def enableRemoveSender(self, enabled):
        self._getButtonRemoveSender().Model.Enabled = enabled

    def removeSender(self, position):
        control = self._getSendersList()
        control.removeItems(position, 1)
        if control.getItemCount() != 0:
            control.selectItemPos(0, True)

    def enableAddRecipient(self, enabled):
        self._getButtonAddRecipient().Model.Enabled = enabled

    def enableRemoveRecipient(self, enabled):
        self._getButtonRemoveRecipient().Model.Enabled = enabled

    def addToRecipient(self):
        control = self._getRecipientsList()
        email = control.getText()
        self.addRecipient(control, email)

    def addRecipient(self, control, email):
        control.setText('')
        count = control.getItemCount()
        control.addItem(email, count)
        self._getButtonRemoveRecipient().Model.Enabled = False

    def removeRecipient(self):
        self._getButtonRemoveRecipient().Model.Enabled = False
        control = self._getRecipientsList()
        email = control.getText()
        recipients = control.getItems()
        if email in recipients:
            control.setText('')
            position = recipients.index(email)
            control.removeItems(position, 1)

    def enableButtonSend(self, enabled):
        self._getButtonSend().Model.Enabled = enabled

    def setStep(self, step):
        self._page1.Model.Step = step

    def isDisposed(self):
        return self._window is None

    def dispose(self):
        self._dialog.dispose()
        self._window.dispose()
        self._page1.dispose()
        self._page2.dispose()
        self._dialog = None
        self._window = None
        self._page1 = None
        self._page2 = None

# MailerView getter methods
    def execute(self):
        return self._dialog.execute()

    def getSender(self):
        control = self._getSendersList()
        sender = control.getSelectedItem()
        position = control.getSelectedItemPos()
        return sender, position

# MailerView private setter methods


# MailerView private getter methods
    def _getTabPageTitle(self, model, id):
        resource = self._getTabResource(id)
        return model.resolveString(resource)

# MailerView StringRessoure methods
    def getTitleRessource(self):
        return 'MailerDialog.Title'

    def getPropertyResource(self, name):
        return 'Mailer.Document.Property.%s' % name

    def getDocumentResource(self):
        return 'MailerPage1.Label2.Label'

    def _getTabResource(self, id):
        return 'MailerWindow.Tab1.Page%s.Title' % id

# MailerView private control methods
    def _getSendersList(self):
        return self._window.getControl('ListBox1')

    def _getButtonRemoveSender(self):
        return self._window.getControl('CommandButton2')

    def _getButtonAddRecipient(self):
        return self._window.getControl('CommandButton3')

    def _getButtonRemoveRecipient(self):
        return self._window.getControl('CommandButton4')

    def _getRecipientsList(self):
        return self._window.getControl('ComboBox1')

    def _getSaveSubject(self):
        return self._window.getControl('CheckBox1')

    def _getSubject(self):
        return self._window.getControl('TextField1')

    def _getSendAsHtml(self):
        return self._page1.getControl('OptionButton1')

    def _getSendAsAttachment(self):
        return self._page1.getControl('OptionButton2')

    def _getDocumentLabel(self):
        return self._page1.getControl('Label2')

    def _getButtonViewHtml(self):
        return self._page1.getControl('CommandButton1')

    def _getSaveMessage(self):
        return self._page1.getControl('CheckBox1')

    def _getAttachMode(self):
        return self._page1.getControl('CheckBox2')

    def _getMessage(self):
        return self._page1.getControl('TextField1')

    def _getSaveAttachments(self):
        return self._page2.getControl('CheckBox2')

    def _getAttachments(self):
        return self._page2.getControl('ListBox1')

    def _getButtonSend(self):
        return self._dialog.getControl('CommandButton2')

# MailerView private methods
    def _getTabPages(self, name, point, size, title1, title2, id):
        service = 'com.sun.star.awt.tab.UnoControlTabPageContainerModel'
        model = self._window.Model.createInstance(service)
        model.PositionX = point.X
        model.PositionY = point.Y
        model.Width = size.Width
        model.Height = size.Height
        self._window.Model.insertByName(name, model)
        tab = self._window.getControl(name)
        page1 = self._getTabPage(tab, model, title1, 0)
        page2 = self._getTabPage(tab, model, title2, 1)
        tab.ActiveTabPageID = id
        return page1, page2

    def _getTabPage(self, tab, model, title, id):
        page = model.createTabPage(id + 1)
        page.Title = title
        index = model.getCount()
        model.insertByIndex(index, page)
        return tab.getControls()[id]
