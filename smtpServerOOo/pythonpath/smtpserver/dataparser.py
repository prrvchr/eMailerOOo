#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.auth import XRestDataParser

from unolib import getNamedValue
from unolib import KeyMap

import xml.etree.ElementTree as XmlTree


class DataParser(unohelper.Base,
                 XRestDataParser):
    def __init__(self):
        self.DataType = 'Xml'

    def parseResponse(self, response):
        config = KeyMap()
        provider = XmlTree.fromstring(response).find('emailProvider')
        config.insertValue('Provider', provider.attrib['id'])
        displayname = provider.find('displayName').text
        config.insertValue('DisplayName', displayname)
        displayshortname = provider.find('displayShortName').text
        config.insertValue('DisplayShortName', displayshortname)
        domains = []
        for domain in provider.findall('domain'):
            domains.append(domain.text)
        config.insertValue('Domains', tuple(domains))
        servers = []
        for s in provider.findall('outgoingServer'):
            server = KeyMap()
            server.insertValue('Server', s.find('hostname').text)
            server.insertValue('Port', s.find('port').text)
            server.insertValue('Connection', self._getConnexion(s))
            server.insertValue('UserName', self._getUserName(s))
            server.insertValue('Authentication', self._getAuthentication(s))
            servers.append(server)
        config.insertValue('Servers', tuple(servers))
        print("DataParser.parseResponse()\n%s" % response)
        print("DataParser.parseResponse()\n%s" % config)
        return config

    def _getConnexion(self, server):
        map = {'plain': 0, 'SSL': 1, 'STARTTLS': 2}
        return map.get(server.find('socketType').text, 0)

    def _getAuthentication(self, server):
        map = {'none': 0, 'password-cleartext': 1, 'plain': 1, 'password-encrypted': 2, 
               'secure': 2, 'OAuth2': 3}
        return map.get(server.find('authentication').text, 0)

    def _getUserName(self, server):
        map = {'%EMAILADDRESS%': -1, '%EMAILLOCALPART%': 0, '%EMAILDOMAIN%': 1}
        return map.get(server.find('username').text, -1)
