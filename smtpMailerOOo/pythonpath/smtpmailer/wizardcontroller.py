#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.ui.dialogs import XWizardController
from com.sun.star.sdb.CommandType import COMMAND
from com.sun.star.sdb.CommandType import QUERY
from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.awt import XCallback
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK
from com.sun.star.ui.dialogs.ExecutableDialogResults import CANCEL
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import PropertySet
from unolib import createService
from unolib import getConfiguration
from unolib import getStringResource
from unolib import getContainerWindow
from unolib import getDialogUrl

from .configuration import g_identifier
from .configuration import g_extension
from .configuration import g_column_index
from .configuration import g_column_filters

from .wizardhandler import WizardHandler
from .wizardpage import WizardPage
from .dbqueries import getSqlQuery
from .logger import logMessage

import traceback

MOTOBIT = 1024 * 1024

class WizardController(unohelper.Base,
                       XWizardController):
    def __init__(self, ctx, wizard):
        self.ctx = ctx
        self._wizard = wizard
        self._pages = [1]
        self._provider = createService(self.ctx, 'com.sun.star.awt.ContainerWindowProvider')
        self._stringResource = getStringResource(self.ctx, g_identifier, g_extension)
        self._configuration = getConfiguration(self.ctx, g_identifier, True)
        self._handler = WizardHandler(self.ctx, self._wizard)
        self._maxsize = self._configuration.getByName("MaxSizeMo") * MOTOBIT

    # XWizardController
    def createPage(self, parent, id):
        try:
            msg = "PageId: %s ..." % id
            url = getDialogUrl(g_extension, 'PageWizard%s' % id)
            window = self._provider.createContainerWindow(url, '', parent, self._handler)
            page = WizardPage(self.ctx, id, window, self._handler)
            msg += " Done"
            logMessage(self.ctx, INFO, msg, 'WizardController', 'createPage()')
            return page
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            logMessage(self.ctx, SEVERE, msg, 'WizardController', 'createPage()')

    def getPageTitle(self, id):
        title = self._stringResource.resolveString('PageWizard%s.Step' % (id, ))
        return title

    def canAdvance(self):
        return self._wizard.CurrentPage.canAdvance()

    def onActivatePage(self, id):
        try:
            msg = "PageId: %s..." % id
            self._wizard.setTitle(self.getPageTitle(id))
            backward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.PREVIOUS')
            forward = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.NEXT')
            finish = uno.getConstantByName('com.sun.star.ui.dialogs.WizardButton.FINISH')
            self._wizard.enableButton(finish, False)
            if self._isFirstLoad(id) and self.canAdvance():
                self._wizard.travelNext()
            self._wizard.updateTravelUI()
            msg += " Done"
            logMessage(self.ctx, INFO, msg, 'WizardController', 'onActivatePage()')
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            logMessage(self.ctx, SEVERE, msg, 'WizardController', 'onActivatePage()')

    def onDeactivatePage(self, id):
        try:
            if id == 1:
                pass
            elif id == 2:
                pass
            elif id == 3:
                pass
        except Exception as e:
            msg = "Error: %s - %s" % (e, traceback.print_exc())
            logMessage(self.ctx, SEVERE, msg, 'WizardController', 'onDeactivatePage()')

    def confirmFinish(self):
        return True

    def _isFirstLoad(self, id):
        if id in self._pages:
            self._pages.remove(id)
            return True
        return False
