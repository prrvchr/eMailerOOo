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

from com.sun.star.mail.MailServiceType import IMAP
from com.sun.star.mail.MailServiceType import SMTP

from com.sun.star.ucb.ConnectionMode import OFFLINE

from com.sun.star.rest.HTTPStatusCode import NOT_FOUND

from .ispdbserver import IspdbServer

from ...runner import RunnerException
from ...runner import Task

from ...unotool import Resource

from ...user import User

from ...oauth20 import getRequest

from ...unotool import getConfiguration
from ...unotool import getConnectionMode
from ...unotool import getUrl

from ...configuration import g_chunk
from ...configuration import g_identifier

import xml.etree.ElementTree as ET
import traceback


class Ispdb(Task):
    def __init__(self, ctx, callback, sender=''):
        super().__init__(ctx, 'Ispdb', callback)
        self._config = None
        self._url = None
        self._user = None
        self._mode = OFFLINE
        self._request = None
        self._response = None
        self._servers = None
        self._resources = {1: Resource(10101, lambda x: (x.total, )),
                           2: Resource(10102, lambda x: (x.total, )),
                           3: Resource(10103),
                           4: Resource(10104),
                           5: Resource(10105),
                           6: Resource(10106)}
        self.total = 6
        self.steps = self._getJobSteps(ctx, sender)

    def callback(self, success=True, error=None):
        self._callback(success, self._servers, self._user, self._mode, error)

    def stepInitJob(self, ctx, sender):
        self._config = getConfiguration(ctx, g_identifier)
        self._url = getUrl(ctx, self._config.getByName('IspDBUrl'))
        self._user = User(ctx, sender)
        self._servers = IspdbServer(self._user)

    def stepGetConnectionMode(self, ctx):
        self._mode = getConnectionMode(ctx, self._url.Server)

    def stepGetHttpRequest(self, ctx):
        self._request = getRequest(ctx)

    def stepGetIspdbConfig(self):
        try:
            parameter = self._request.getRequestParameter('getIspdbConfig')
            parameter.Url = self._url.Complete + self._user.getUserDomain()
            parameter.NoAuth = True
            response = self._request.execute(parameter)
            if response.Ok or response.StatusCode == NOT_FOUND:
                self._response = response
            else:
                response.raiseForStatus()
        except Exception as e:
            trace = traceback.format_exc()
            raise RunnerException(e, trace)

    def stepParseIspdbConfig(self):
        smtps = []
        imaps = []
        config = {}
        provider = None
        map1 = {'none': 0, 'password-cleartext': 1, 'plain': 1,
                'password-encrypted': 1, 'secure': 1, 'OAuth2': 2}
        map2 = {'plain': 0, 'SSL': 1, 'STARTTLS': 2}
        map3 = {'%EMAILLOCALPART%': 0, '%EMAILADDRESS%': 1, '%EMAILDOMAIN%': 2}
        parser = ET.XMLPullParser(('end', ))
        iterator = self._response.iterContent(g_chunk, False)
        while iterator.hasMoreElements():
            # FIXME: As Decode is False we obtain a sequence of bytes
            parser.feed(iterator.nextElement().value)
            for _, element in parser.read_events():
                if element.tag == 'emailProvider':
                    provider = element.get('id')
                    for child in element.findall('outgoingServer'):
                        if child.get('type') == 'smtp':
                            smtps.append(self._parseServer(provider, child, map1, map2, map3))
                    for child in element.findall('incomingServer'):
                        if child.get('type') == 'imap':
                            imaps.append(self._parseServer(provider, child, map1, map2, map3))
        config[SMTP.value] = tuple(smtps)
        config[IMAP.value] = tuple(imaps)
        if self._user.isNew():
            self._user.enableImap(len(imaps) > 0)
            clients = self._config.getByName('Urls').getElementNames()
            if provider in clients:
                self._user.setClient(provider)
        # XXX: If provider use API and we need to be able to test the connection
        # XXX: then it is necessary to save the domain used by the user
        if provider:
            self._user.setProvider(provider)
            self._user.setAppsLink(self._getAppsLinks().get(provider, ''))
        self._servers = IspdbServer(self._user, config)

    def stepFinalize(self):
        if self._response:
            self._response.close()
        self._response = None
        if self._request:
            self._request.dispose()
        self._request = None

    def _getStepInitJob(self, ctx, sender, name):
        return 10201, (name, ), self.stepInitJob, ctx, sender

    def _getStepGetConnectionMode(self, ctx):
        return 10301, (), self.stepGetConnectionMode, ctx

    def _getStepGetHttpRequest(self, ctx):
        return 10401, (), self.stepGetHttpRequest, ctx

    def _getStepGetIspdbConfig(self):
        return 10501, (), self.stepGetIspdbConfig

    def _getStepParseIspdbConfig(self):
        return 10601, (), self.stepParseIspdbConfig

    def _getStepFinalize(self, offset=0):
        return 10701 + offset, (), self.stepFinalize

    def _getJobSteps(self, ctx, sender):
        offset = 0
        yield self._getStepInitJob(ctx, sender, self.__class__.__name__)
        yield self._getStepGetConnectionMode(ctx)
        if self._mode == OFFLINE:
            offset = 1 if self._user.isNew() else 2
            yield self._getStepSkipStep(10301)
            yield self._getStepSkipStep(10301)
            yield self._getStepSkipStep(10301)
        else:
            yield self._getStepGetHttpRequest(ctx)
            if self._request is None:
                offset = 3 if self._user.isNew() else 4
                yield self._getStepSkipStep(10401)
                yield self._getStepSkipStep(10401)
            else:
                yield self._getStepGetIspdbConfig()
                if self._response.StatusCode == NOT_FOUND:
                    offset = 5
                    yield self._getStepSkipStep(10501)
                else:
                    yield self._getStepParseIspdbConfig()
                    offset = 6
        yield self._getStepFinalize(offset)

    def _parseServer(self, provider, element, map1, map2, map3):
        hostname = element.find('hostname').text
        authentication = map1.get(element.find('authentication').text, 1)
        if authentication == 2 and not self._supportOAuth2(provider, hostname):
            authentication = 1
        server = {'ServerName': hostname,
                  'Port': int(element.find('port').text),
                  'AuthenticationType': authentication,
                  'ConnectionType': map2.get(element.find('socketType').text, 1),
                  'LoginMode': map3.get(element.find('username').text, 1)}
        return server

    def _supportOAuth2(self, provider, hostname):
        # FIXME: We support OAuth2 protocol only for providers:
        # FIXME: - Who have registered their hostname as a ResourceURL
        # FIXME: - Which uses an API to replace SMTP and/or IMAP servers
        providers = self._config.getByName('Providers')
        return providers.hasByName(provider) or self._request.isRegisteredUrl(hostname)

    def _getAppsLinks(self):
        links = {}
        applications = self._config.getByName('Applications')
        for name in applications.getElementNames():
            apps = applications.getByName(name)
            links[name] = apps.getByName('Link')
        return links
