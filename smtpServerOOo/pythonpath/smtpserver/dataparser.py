#!
# -*- coding: utf_8 -*-

'''
    Copyright (c) 2020 https://prrvchr.github.io

    Permission is hereby granted, free of charge, to any person obtaining
    a copy of this software and associated documentation files (the "Software"),
    to deal in the Software without restriction, including without limitation
    the rights to use, copy, modify, merge, publish, distribute, sublicense,
    and/or sell copies of the Software, and to permit persons to whom the Software
    is furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
    EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
    OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
    CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
    TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
    OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'''

import uno
import unohelper

from com.sun.star.auth import XRestDataParser

from unolib import KeyMap

import xml.etree.ElementTree as XmlTree


class DataParser(unohelper.Base,
                 XRestDataParser):
    def __init__(self):
        self.DataType = 'Xml'

    def parseResponse(self, response):
        data = KeyMap()
        config = XmlTree.fromstring(response).find('emailProvider')
        provider = config.attrib['id']
        data.insertValue('Provider', provider)
        displayname = config.find('displayName').text
        data.insertValue('DisplayName', displayname)
        displayshortname = config.find('displayShortName').text
        data.insertValue('DisplayShortName', displayshortname)
        domains = []
        for domain in config.findall('domain'):
            if domain.text != provider:
                domains.append(domain.text)
        data.insertValue('Domains', tuple(domains))
        servers = []
        for s in config.findall('outgoingServer'):
            server = KeyMap()
            server.insertValue('Server', s.find('hostname').text)
            server.insertValue('Port', self._getPort(s))
            server.insertValue('Connection', self._getConnexion(s))
            server.insertValue('Authentication', self._getAuthentication(s))
            server.insertValue('LoginMode', self._getLoginMode(s))
            servers.append(server)
        data.insertValue('Servers', tuple(servers))
        return data

    def _getPort(self, server):
        return int(server.find('port').text)

    def _getConnexion(self, server):
        map = {'plain': 0, 'SSL': 1, 'STARTTLS': 2}
        return map.get(server.find('socketType').text, 0)

    def _getAuthentication(self, server):
        map = {'none': 0, 'password-cleartext': 1, 'plain': 1, 'password-encrypted': 2, 
               'secure': 2, 'OAuth2': 3}
        return map.get(server.find('authentication').text, 0)

    def _getLoginMode(self, server):
        map = {'%EMAILLOCALPART%': 0, '%EMAILADDRESS%': 1, '%EMAILDOMAIN%': 2}
        return map.get(server.find('username').text, 1)
