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

from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.frame import XDispatchResultListener

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import getFileSequence
from unolib import getStringResource
from unolib import getResourceLocation
from unolib import getUrl
from unolib import getDialog
from unolib import createService
from unolib import getPropertyValueSet
from unolib import executeDispatch

from smtpserver import PageModel

from smtpserver import getLoggerUrl
from smtpserver import getLoggerSetting
from smtpserver import setLoggerSetting
from smtpserver import clearLogger
from smtpserver import logMessage
from smtpserver import getMessage
g_message = 'OptionsDialog'

from smtpserver import g_extension
from smtpserver import g_identifier

import os
import sys
import traceback


# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = '%s.OptionsDialog' % g_identifier


class OptionsDialog(unohelper.Base,
                    XServiceInfo,
                    XContainerWindowEventHandler,
                    XDialogEventHandler,
                    XDispatchResultListener):
    def __init__(self, ctx):
        try:
            self._ctx = ctx
            self._stringResource = getStringResource(self._ctx, g_identifier, g_extension, 'OptionsDialog')
            self._spooler = createService(self._ctx, 'com.sun.star.mail.MailServiceSpooler')
            self._model = PageModel(self._ctx)
            logMessage(self._ctx, INFO, "Loading ... Done", 'OptionsDialog', '__init__()')
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    # XContainerWindowEventHandler, XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        handled = False
        if method == 'external_event':
            print("OptionsDialog.callHandlerMethod() %s" % event)
            if event == 'ok':
                self._saveSetting(dialog)
                handled = True
            elif event == 'back':
                self._loadSetting(dialog)
                handled = True
            elif event == 'initialize':
                self._loadSetting(dialog)
                handled = True
        elif method == 'ToggleLogger':
            enabled = event.Source.State == 1
            self._toggleLogger(dialog, enabled)
            handled = True
        elif method == 'EnableViewer':
            self._toggleViewer(dialog, True)
            handled = True
        elif method == 'DisableViewer':
            self._toggleViewer(dialog, False)
            handled = True
        elif method == 'ViewLog':
            self._viewLog(dialog)
            handled = True
        elif method == 'ClearLog':
            self._clearLog(dialog)
            handled = True
        elif method == 'LogInfo':
            self._logInfo(dialog)
            handled = True
        elif method == 'ChangeTimeout':
            self._changeTimeout(event.Source)
            handled = True
        elif method == 'ShowWizard':
            self._showSmtpServer(dialog)
            handled = True
        elif method == 'ToggleSpooler':
            self._toogleSpooler(dialog)
            handled = True
        elif method == 'ShowSpooler':
            self._showSmtpSpooler(dialog)
            handled = True
        return handled
    def getSupportedMethodNames(self):
        return ('external_event', 'ToggleLogger', 'EnableViewer', 'DisableViewer',
                'ViewLog', 'ClearLog', 'LogInfo', 'ChangeTimeout', 'ShowWizard',
                'ToggleSpooler', 'ShowSpooler')

    def _loadSetting(self, dialog):
        self._loadLoggerSetting(dialog)
        self._loadSmtpSetting(dialog)
        self._loadSpooler(dialog)

    def _saveSetting(self, dialog):
        self._saveLoggerSetting(dialog)
        self._saveSmtpSetting(dialog)

    def _toggleLogger(self, dialog, enabled):
        dialog.getControl('Label1').Model.Enabled = enabled
        dialog.getControl('ListBox1').Model.Enabled = enabled
        dialog.getControl('OptionButton1').Model.Enabled = enabled
        control = dialog.getControl('OptionButton2')
        control.Model.Enabled = enabled
        self._toggleViewer(dialog, enabled and control.State)

    def _toggleViewer(self, dialog, enabled):
        dialog.getControl('CommandButton1').Model.Enabled = enabled

    def _viewLog(self, window):
        dialog = getDialog(self._ctx, g_extension, 'LogDialog', self, window.Peer)
        url = getLoggerUrl(self._ctx)
        dialog.Title = url
        self._setDialogText(dialog, url)
        dialog.execute()
        dialog.dispose()

    def _clearLog(self, dialog):
        try:
            clearLogger()
            msg = getMessage(self._ctx, g_message, 101)
            logMessage(self._ctx, INFO, msg, 'OptionsDialog', '_clearLog()')
            url = getLoggerUrl(self._ctx)
            self._setDialogText(dialog, url)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            logMessage(self._ctx, SEVERE, msg, "OptionsDialog", "_clearLog()")

    def _logInfo(self, dialog):
        version  = ' '.join(sys.version.split())
        msg = getMessage(self._ctx, g_message, 111, version)
        logMessage(self._ctx, INFO, msg, "OptionsDialog", "_logInfo()")
        path = os.pathsep.join(sys.path)
        msg = getMessage(self._ctx, g_message, 112, path)
        logMessage(self._ctx, INFO, msg, "OptionsDialog", "_logInfo()")
        url = getLoggerUrl(self._ctx)
        self._setDialogText(dialog, url)

    def _setDialogText(self, dialog, url):
        control = dialog.getControl('TextField1')
        length, sequence = getFileSequence(self._ctx, url)
        control.Text = sequence.value.decode('utf-8')
        selection = uno.createUnoStruct('com.sun.star.awt.Selection', length, length)
        control.setSelection(selection)

    def _loadLoggerSetting(self, dialog):
        enabled, index, handler = getLoggerSetting(self._ctx)
        dialog.getControl('CheckBox1').State = int(enabled)
        dialog.getControl('ListBox1').selectItemPos(index, True)
        dialog.getControl('OptionButton%s' % handler).State = 1
        self._toggleLogger(dialog, enabled)

    def _saveLoggerSetting(self, dialog):
        enabled = bool(dialog.getControl('CheckBox1').State)
        index = dialog.getControl('ListBox1').getSelectedItemPos()
        handler = dialog.getControl('OptionButton1').State
        setLoggerSetting(self._ctx, enabled, index, handler)

    def _loadSmtpSetting(self, dialog):
        dialog.getControl('NumericField1').Value = self._model.Timeout

    def _saveSmtpSetting(self, dialog):
        self._model.saveTimeout()

    def _loadSpooler(self, dialog):
        resource = 'OptionsDialog.Label5.Label.%s' % int(self._spooler.isStarted())
        dialog.getControl('Label5').Text = self._stringResource.resolveString(resource)

    def _changeTimeout(self, control):
        self._model.Timeout = int(control.Value)

    def _showSmtpServer(self, dialog):
        try:
            executeDispatch(self._ctx, 'smtp:server')
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def _toogleSpooler(self, dialog):
        if self._spooler.isStarted():
            self._spooler.stop()
        else:
            self._spooler.start()
        self._loadSpooler(dialog)

    def _showSmtpSpooler(self, dialog):
        try:
            print("OptionsDialog._showSpooler() 1")
            #self._spooler.viewSpooler(dialog.Peer)
            executeDispatch(self._ctx, 'smtp:spooler')
            print("OptionsDialog._showSpooler() 2")
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    # XDispatchResultListener
    def dispatchFinished(self, notification):
        print("OptionsDialog.dispatchFinished() %s" % (notification.Result.getKeys(), ))
    def disposing(self, source):
        pass

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)


g_ImplementationHelper.addImplementation(OptionsDialog,                             # UNO object class
                                         g_ImplementationName,                      # Implementation name
                                        (g_ImplementationName,))                    # List of implemented services
