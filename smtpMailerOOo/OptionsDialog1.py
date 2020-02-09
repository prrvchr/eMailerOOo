#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.awt import XContainerWindowEventHandler
from com.sun.star.beans import PropertyValue


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.smtpMailerOOo.OptionsDialog"


class PyOptionsDialog(unohelper.Base, XServiceInfo, XContainerWindowEventHandler):
    def __init__(self, ctx):
        self.ctx = ctx
        self.dialog = None
        self.configuration = "com.gmail.prrvchr.extensions.smtpMailerOOo/Options"
        return

    # XContainerWindowEventHandler
    def callHandlerMethod(self, dialog, event, method):
        if dialog.Model.Name == "OptionsDialog":
            if method == "external_event":
                if event == "ok":
                    self._saveSetting()
                    return True
                elif event == "back":
                    self._loadSetting()
                    return True
                elif event == "initialize":
                    if self.dialog is None:
                        self.dialog = dialog
                    self._loadSetting()
                    return True
        return False
    def getSupportedMethodNames(self):
        return ("external_event",)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _loadSetting(self):
        configuration = self._getConfiguration(self.configuration)
        self.dialog.getControl("SendAsHtml").setState(configuration.getByName("SendAsHtml"))
        self.dialog.getControl("SendAsPdf").setState(configuration.getByName("SendAsPdf"))
        self.dialog.getControl("MaxSizeMo").setValue(configuration.getByName("MaxSizeMo"))
        self.dialog.getControl("OffLineUse").setState(configuration.getByName("OffLineUse"))

    def _saveSetting(self):
        configuration = self._getConfiguration(self.configuration, True)
        configuration.replaceByName("SendAsHtml", self.dialog.getControl("SendAsHtml").getState() != 0)
        configuration.replaceByName("SendAsPdf", self.dialog.getControl("SendAsPdf").getState() != 0)
        configuration.replaceByName("MaxSizeMo", int(self.dialog.getControl("MaxSizeMo").getValue()))
        configuration.replaceByName("OffLineUse", self.dialog.getControl("OffLineUse").getState() != 0)
        configuration.commitChanges()

    def _getConfiguration(self, nodepath, update=False):
        value = uno.Enum("com.sun.star.beans.PropertyState", "DIRECT_VALUE")
        config = self.ctx.ServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
        service = "com.sun.star.configuration.ConfigurationUpdateAccess" if update else "com.sun.star.configuration.ConfigurationAccess"
        return config.createInstanceWithArguments(service, (PropertyValue("nodepath", -1, nodepath, value),))


# uno implementation
g_ImplementationHelper.addImplementation(PyOptionsDialog,                                            # UNO object class
                                         g_ImplementationName,                                       # Implementation name
                                         ())                                                         # List of implemented services
