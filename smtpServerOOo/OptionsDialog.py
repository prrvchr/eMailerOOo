#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import getFileSequence
from unolib import getStringResource
from unolib import getResourceLocation
from unolib import getDialog
from unolib import createService
from unolib import getPropertyValueSet

from smtpserver import Wizard
from smtpserver import WizardModel
from smtpserver import WizardController
from smtpserver import g_wizard_page
from smtpserver import g_wizard_paths

from smtpserver import getLoggerUrl
from smtpserver import getLoggerSetting
from smtpserver import setLoggerSetting
from smtpserver import clearLogger
from smtpserver import logMessage

from smtpserver import g_extension
from smtpserver import g_identifier

import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = '%s.OptionsDialog' % g_identifier


class OptionsDialog(unohelper.Base,
                    XServiceInfo,
                    XContainerWindowEventHandler,
                    XDialogEventHandler):
    def __init__(self, ctx):
        try:
            self.ctx = ctx
            self.stringResource = getStringResource(self.ctx, g_identifier, g_extension, 'OptionsDialog')
            self._model = WizardModel(self.ctx)
            logMessage(self.ctx, INFO, "Loading ... Done", 'OptionsDialog', '__init__()')
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
        elif method == 'ChangeTimeout':
            self._changeTimeout(event.Source)
            handled = True
        elif method == 'ShowWizard':
            self._showWizard(dialog)
            handled = True
        return handled
    def getSupportedMethodNames(self):
        return ('external_event', 'ToggleLogger', 'EnableViewer', 'DisableViewer',
                'ViewLog', 'ClearLog', 'ChangeTimeout', 'ShowWizard')

    def _loadSetting(self, dialog):
        self._loadLoggerSetting(dialog)
        self._loadSmtpSetting(dialog)

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
        dialog = getDialog(self.ctx, g_extension, 'LogDialog', self, window.Peer)
        url = getLoggerUrl(self.ctx)
        dialog.Title = url
        self._setDialogText(dialog, url)
        dialog.execute()
        dialog.dispose()

    def _clearLog(self, dialog):
        try:
            clearLogger()
            logMessage(self.ctx, INFO, "ClearingLog ... Done", 'OptionsDialog', '_doClearLog()')
            url = getLoggerUrl(self.ctx)
            self._setDialogText(dialog, url)
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            logMessage(self.ctx, SEVERE, msg, "OptionsDialog", "_doClearLog()")

    def _setDialogText(self, dialog, url):
        length, sequence = getFileSequence(self.ctx, url)
        dialog.getControl('TextField1').Text = sequence.value.decode('utf-8')

    def _loadLoggerSetting(self, dialog):
        enabled, index, handler = getLoggerSetting(self.ctx)
        dialog.getControl('CheckBox1').State = int(enabled)
        dialog.getControl('ListBox1').selectItemPos(index, True)
        dialog.getControl('OptionButton%s' % handler).State = 1
        self._toggleLogger(dialog, enabled)

    def _saveLoggerSetting(self, dialog):
        enabled = bool(dialog.getControl('CheckBox1').State)
        index = dialog.getControl('ListBox1').getSelectedItemPos()
        handler = dialog.getControl('OptionButton1').State
        setLoggerSetting(self.ctx, enabled, index, handler)

    def _loadSmtpSetting(self, dialog):
        dialog.getControl('NumericField1').Value = self._model.getTimeout()

    def _saveSmtpSetting(self, dialog):
        self._model.saveTimeout()

    def _changeTimeout(self, control):
        self._model.Timeout = int(control.Value)

    def _showWizard(self, dialog):
        try:
            print("_showWizard() 1")
            url = self._getUrl('ispdb://')
            desktop = createService(self.ctx, 'com.sun.star.frame.Desktop')
            dispatcher = desktop.getCurrentFrame().queryDispatch(url, '', 0)
            #dispatcher = createService(self.ctx, 'com.sun.star.frame.DispatchHelper')
            #dispatcher.executeDispatch(desktop.getCurrentFrame(), 'ispdb://', '', 0, ())
            print("_showWizard() 2")
            if dispatcher is not None:
                args = getPropertyValueSet({'Email': 'prrvchr@gmail.com'})
                dispatcher.dispatch(url, args)
                print("_showWizard() 3")
            #mri = createService(self.ctx, 'mytools.Mri')
            #mri.inspect(desktop)
            msg = "OptionsDialog._showWizard()"
            logMessage(self.ctx, INFO, msg, 'OptionsDialog', '_showWizard()')
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            print(msg)

    def _getUrl(self, uri):
        url = uno.createUnoStruct('com.sun.star.util.URL')
        url.Complete = uri
        transformer = createService(self.ctx, 'com.sun.star.util.URLTransformer')
        success, url = transformer.parseStrict(url)
        return url

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
