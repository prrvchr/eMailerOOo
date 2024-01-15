#!
# -*- coding: utf-8 -*-

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

from com.sun.star.awt.FontWeight import BOLD
from com.sun.star.awt.FontWeight import NORMAL

from ...unotool import getContainerWindow

from ...configuration import g_extension

import traceback


class IspdbView(unohelper.Base):
    def __init__(self, ctx, handler, parent, pageid):
        idl = 'IspdbPage%s' % pageid
        self._dialog = getContainerWindow(ctx, parent, None, g_extension, idl)
        self._window = getContainerWindow(ctx, self._dialog.getPeer(), handler, g_extension, 'IspdbPages')
        self._window.setVisible(True)

# IspdbView getter methods
    def getWindow(self):
        return self._dialog

    def getAuthentication(self):
        return self._getAuthentication().getSelectedItemPos()

    def getConnection(self):
        return self._getConnection().getSelectedItemPos()

    def getData(self):
        return self._getHostValue(), self._getPortValue(), self.getAuthentication()

    def getLogin(self):
        return self._getLogin().Text

    def getPasswords(self):
        return self._getPassword().Text, self._getConfirmPwd().Text

    def getConfiguration(self, service):
        host = self._getHostValue()
        port = self._getPortValue()
        connection = self.getConnection()
        authentication = self.getAuthentication()
        server = self._getServer(service, host, port, connection, authentication)
        user = self._getUser(service, host, port, connection, authentication)
        return server, user

    def getServerConfig(self, service):
        host = self._getHostValue()
        port = self._getPortValue()
        connection = self.getConnection()
        authentication = self.getAuthentication()
        server = self._getServer(service, host, port, connection, authentication)
        return server

# IspdbView setter methods
    def setPageLabel(self, text):
        self._getPageLabel().Text = text

    def updatePage(self, config):
        self._enableNavigation(config)
        self._updateServerPage(config)
        self._setConfiguration(config)

    def setSecurityMessage(self, message, level):
        control = self._getSecurityLabel()
        control.Model.FontWeight = BOLD if level < 3 else NORMAL
        control.Text = message

    def enableLogin(self, enabled):
        self._getLoginLabel().Model.Enabled = enabled
        self._getLogin().Model.Enabled = enabled

    def enablePassword(self, enabled):
        self._getPasswordLabel().Model.Enabled = enabled
        self._getPassword().Model.Enabled = enabled
        self._getConfirmPwdLabel().Model.Enabled = enabled
        self._getConfirmPwd().Model.Enabled = enabled

# IspdbView private getter methods
    def _getHostValue(self):
        return self._getHost().Text

    def _getPortValue(self):
        return int(self._getPort().Value)

    def _getServer(self, service, host, port, connection, authentication):
        server = {}
        server['Service'] = service
        server['ServerName'] = host
        server['Port'] = port
        server['ConnectionType'] = connection
        server['AuthenticationType'] = authentication
        return server

    def _getUser(self, service, host, port, connection, authentication):
        user = {}
        user[service + '/ServerName'] = host
        user[service + '/Port'] = port
        user[service + '/ConnectionType'] = connection
        user[service + '/AuthenticationType'] = authentication
        user[service + '/UserName'] = self._getLogin().Text
        user[service + '/Password'] = self._getPassword().Text
        return user

# IspdbView private setter methods
    def _updateServerPage(self, config):
        default = config.get('Default')
        control = self._getServerPage()
        control.Model.FontWeight = BOLD if default else NORMAL
        control.Text = config.get('Page')

    def _setConfiguration(self, config):
        self._getHost().Text = config.get('ServerName')
        self._getPort().Value = config.get('Port')
        self._getConnection().selectItemPos(config.get('ConnectionType'), True)
        self._getAuthentication().selectItemPos(config.get('AuthenticationType'), True)
        self._getLogin().Text = config.get('UserName')
        self._getPassword().Text = config.get('Password')
        self._getConfirmPwd().Text = config.get('Password')

    def _enableNavigation(self, config):
        self._enablePrevious(config.get('First'))
        self._enableNext(config.get('Last'))

    def _enablePrevious(self, isfirst):
        self._getPrevious().Model.Enabled = not isfirst

    def _enableNext(self, islast):
        self._getNext().Model.Enabled = not islast

# IspdbView private getter control methods
    def _getPageLabel(self):
        return self._window.getControl('Label1')

    def _getServerPage(self):
        return self._window.getControl('Label2')

    def _getHost(self):
        return self._window.getControl('TextField1')

    def _getPort(self):
        return self._window.getControl('NumericField1')

    def _getConnection(self):
        return self._window.getControl('ListBox1')

    def _getAuthentication(self):
        return self._window.getControl('ListBox2')

    def _getLoginLabel(self):
        return self._window.getControl('Label7')

    def _getLogin(self):
        return self._window.getControl('TextField2')

    def _getPasswordLabel(self):
        return self._window.getControl('Label8')

    def _getPassword(self):
        return self._window.getControl('TextField3')

    def _getConfirmPwdLabel(self):
        return self._window.getControl('Label9')

    def _getConfirmPwd(self):
        return self._window.getControl('TextField4')

    def _getPrevious(self):
        return self._window.getControl('CommandButton1')

    def _getNext(self):
        return self._window.getControl('CommandButton2')

    def _getSecurityLabel(self):
        return self._window.getControl('Label10')
