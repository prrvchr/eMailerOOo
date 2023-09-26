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

import uno
import unohelper

from com.sun.star.awt.FontWeight import BOLD

from com.sun.star.ucb.ConnectionMode import OFFLINE

from com.sun.star.logging.LogLevel import ALL

from com.sun.star.mail.MailServiceType import IMAP
from com.sun.star.mail.MailServiceType import SMTP

from com.sun.star.rest.HTTPStatusCode import NOT_FOUND
from com.sun.star.rest import HTTPException
from com.sun.star.rest import RequestException

from com.sun.star.uno import Exception as UnoException

from ..mailerlib import CurrentContext
from ..mailerlib import Authenticator
from ..mailerlib import MailTransferable

from ..unotool import createService
from ..unotool import getConfiguration
from ..unotool import getConnectionMode
from ..unotool import getResourceLocation
from ..unotool import getStringResource
from ..unotool import getUrl

from ..mailertool import getMail
from ..mailertool import getMessageImage

from ..logger import LogController
from ..logger import RollerHandler

from ..oauth2 import g_oauth2

from ..configuration import g_identifier
from ..configuration import g_extension
from ..configuration import g_mailservicelog
from ..configuration import g_logo
from ..configuration import g_logourl
from ..configuration import g_chunk

from .pages import IspdbServer
from .pages import IspdbUser

import xml.etree.ElementTree as ET
from threading import Thread
import validators
import traceback


class IspdbModel(unohelper.Base):
    def __init__(self, ctx, datasource, close, email=''):
        self._ctx = ctx
        self._datasource = datasource
        self._close = close
        self._email = email
        self._user = None
        self._servers = None
        self._services = [SMTP.value, IMAP.value]
        self._count = 0
        self._offline = 0
        self._isoauth2 = True
        self._isnew = False
        self._diposed = False
        self._updated = False
        self._version = 0
        self._messageid = None
        self._listener = None
        secure = {0: 3, 1: 4, 2: 4, 3: 5}
        unsecure = {0: 0, 1: 1, 2: 2, 3: 2}
        self._levels = {0: unsecure, 1: secure, 2: secure}
        self._connections = {0: 'Insecure', 1: 'Ssl', 2: 'Tls'}
        self._authentications = {0: 'None', 1: 'Login', 2: 'Login', 3: 'OAuth2'}
        configuration = getConfiguration(ctx, g_identifier)
        self._url = configuration.getByName('IspDBUrl')
        self._timeout = configuration.getByName('ConnectTimeout')
        self._logger = LogController(ctx, g_mailservicelog)
        self._resolver = getStringResource(ctx, g_identifier, g_extension)
        self._resources = {'Step': 'IspdbPage%s.Step',
                           'Title': 'IspdbPage%s.Title.%s',
                           'PageLabel': 'IspdbPage%s.Label1.Label',
                           'PagesLabel': 'IspdbPages.Label1.Label',
                           'Progress': 'IspdbPage2.Label2.Label.%s',
                           'Security': 'IspdbPages.Label10.Label.%s',
                           'SendTitle': 'SendDialog.Title',
                           'SendSubject': 'SendDialog.TextField2.Text',
                           'SendMessage': 'SendDialog.TextField3.Text',
                           'ThreadTitle': 'SendThread.Title'}

    @property
    def Email(self):
        return self._email
    @Email.setter
    def Email(self, email):
        self._email = email

    @property
    def Offline(self):
        return self._offline
    @Offline.setter
    def Offline(self, offline):
        self._offline = offline

    @property
    def DataSource(self):
        return self._datasource

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

# IspdbModel getter methods called by IspdbWizard
    def dispose(self):
        self._diposed = True
        if self._listener is not None:
            self._logger.removeModifyListener(self._listener)
            self._listener = None
        if self._close:
            self.DataSource.dispose()

# IspdbModel getter methods called by WizardPages 2 and 4
    def isDisposed(self):
        return self._diposed

# IspdbModel getter methods called by WizardPage1
    def isEmailValid(self, email):
        if validators.email(email):
            return True
        return False

# IspdbModel getter methods called by WizardPage2
    def getServerConfig(self, *args):
        self._version += 1
        Thread(target=self._getServerConfig, args=args).start()

    def _getServerConfig(self, progress, updateModel):
        servers = IspdbServer()
        progress(5)
        url = getUrl(self._ctx, self._url)
        progress(10)
        mode = getConnectionMode(self._ctx, url.Server)
        progress(20)
        self.DataSource.waitForDataBase()
        progress(40)
        user = IspdbUser(self.DataSource.getServerConfig(servers, self.Email))
        if servers.hasServers(self._services):
            progress(100, 1)
        elif mode == OFFLINE:
            progress(100, 2, BOLD)
        else:
            progress(60)
            request = createService(self._ctx, g_oauth2)
            if request is None:
                self._isoauth2 = False
                progress(100, 3, BOLD)
            else:
                progress(70)
                try:
                    config = self._getIspdbConfig(request, url.Complete, user.getDomain())
                except RequestException as e:
                    progress(100, 4, BOLD)
                else:
                    progress(80)
                    if config is None:
                        progress(100, 5)
                    else:
                        self.DataSource.setServerConfig(self._services, servers, config)
                        progress(100, 6)
        updateModel(user, servers, mode)

    def _getIspdbConfig(self, request, url, domain):
        config = None
        parameter = request.getRequestParameter('getIspdbConfig')
        parameter.Url = '%s%s' % (url, domain)
        parameter.NoAuth = True
        response = request.execute(parameter)
        if response.Ok:
            config = self._parseIspdbConfig(response)
        elif response.StatusCode != NOT_FOUND:
            response.raiseForStatus()
        response.close()
        return config

    def _parseIspdbConfig(self, response):
        smtps = []
        imaps = []
        domains = []
        config = {}
        map1 = {'plain': 0, 'SSL': 1, 'STARTTLS': 2}
        map2 = {'none': 0, 'password-cleartext': 1, 'plain': 1,
                'password-encrypted': 2, 'secure': 2, 'OAuth2': 3}
        map3 = {'%EMAILLOCALPART%': 0, '%EMAILADDRESS%': 1, '%EMAILDOMAIN%': 2}
        parser = ET.XMLPullParser(('end', ))
        iterator = response.iterContent(g_chunk, False)
        while iterator.hasMoreElements():
            # FIXME: As Decode is False we obtain a sequence of bytes
            parser.feed(iterator.nextElement().value)
            for event, element in parser.read_events():
                if element.tag != 'emailProvider':
                    continue
                provider = element.get('id')
                config['Provider'] = provider
                config['DisplayName'] = element.find('displayName').text
                config['DisplayShortName'] = element.find('displayShortName').text
                for child in element.findall('domain'):
                    if child.text != provider:
                        domains.append(child.text)
                for child in element.findall('outgoingServer'):
                    if child.get('type') == 'smtp':
                        smtps.append(self._parseServer(child, SMTP.value, map1, map2, map3))
                for child in element.findall('incomingServer'):
                    if child.get('type') == 'imap':
                        imaps.append(self._parseServer(child, IMAP.value, map1, map2, map3))
        config['Domains'] = tuple(domains)
        config[SMTP.value] = tuple(smtps)
        config[IMAP.value] = tuple(imaps)
        return config

    def _parseServer(self, element, service, map1, map2, map3):
        server = {}
        server['Service'] = service
        server['Server'] = element.find('hostname').text
        server['Port'] = int(element.find('port').text)
        server['Connection'] = map1.get(element.find('socketType').text, 0)
        server['Authentication'] = map2.get(element.find('authentication').text, 0)
        server['LoginMode'] = map3.get(element.find('username').text, 1)
        return server

    def setServerConfig(self, user, servers, offline):
        self._user = user
        self._servers = servers.setDefault(self._services, user)
        self._offline = offline

# IspdbModel getter methods called by WizardPage3 / WizardPage4
    def isOAuth2(self):
        return self._isoauth2

    def refreshView(self, version):
        return version < self._version

    def getVersion(self):
        return self._version

    def getConfig(self, service):
        config = {}
        config['First'] = self._servers.isFirst(service)
        config['Last'] = self._servers.isLast(service)
        config['Page'] = self._servers.getServerPage(service)
        config['Default'] = self._servers.isDefaultPage(service, self._user)
        config.update(self._servers.getCurrentServer(service))
        config['Login'] = self._getLogin(service)
        config['Password'] = self._getPassword(service)
        return config

    def isConnectionValid(self, host, port):
        return self._isHostValid(host) and self._isPortValid(port)

    def isStringValid(self, value):
        if validators.length(value, min=1):
            return True
        return False

    def previousServerPage(self, service, server, user):
        self.updateConfiguration(service, server, user)
        self._servers.decrementIndex(service)

    def nextServerPage(self, service, server, user):
        self.updateConfiguration(service, server, user)
        self._servers.incrementIndex(service)

    def updateConfiguration(self, service, server, user):
        self._servers.updateCurrentServer(service, server)
        self._user.updateUser(user)

    def saveConfiguration(self):
        provider = self._user.getDomain()
        for service in self._services:
            self._servers.saveServer(self._datasource, service, provider)
        if self._user.isUpdated():
            self._datasource.saveUser(self.Email, self._user)

    def getSecurity(self, i, j):
        level = self._levels.get(i).get(j)
        message = self.getSecurityMessage(level)
        return message, level

# IspdbModel getter methods called by WizardPage5
    def addLogListener(self, listener):
        self._logger.addModifyListener(listener)
        self._listener = listener

    def getLogContent(self):
        return self._logger.getLogContent(True)

    def connectServers(self, *args):
        Thread(target=self._connectServers, args=args).start()

    def _connectServers(self, reset, progress, setlabel, setstep):
        handler = RollerHandler(self._ctx, self._logger.Name)
        self._logger.addRollerHandler(handler)
        i = 0
        step = 2
        range = 100
        reset(self._getServicesCount() * range)
        for service in self._services:
            setlabel(service, self._getHost(service), self._getPort(service))
            context = self._getConnectionContext(service)
            authenticator = self._getAuthenticator(service)
            stype = uno.Enum('com.sun.star.mail.MailServiceType', service)
            step = self._connectServer(context, authenticator, stype, i * range, progress)
            i += 1
        self._logger.removeRollerHandler(handler)
        setstep(step)

    def _connectServer(self, context, authenticator, stype, i, progress):
        step = 2
        progress(i + 5)
        service = 'com.sun.star.mail.MailServiceProvider2'
        progress(i + 25)
        host = context.getValueByName('ServerName')
        server = createService(self._ctx, service).create(stype, host)
        progress(i + 50)
        try:
            server.connect(context, authenticator)
        except UnoException as e:
            progress(i + 100)
        else:
            progress(i + 75)
            if server.isConnected():
                server.disconnect()
                step = 4
                progress(i + 100)
            else:
                progress(i + 100)
        return step

    def _getHost(self, service):
        return self._servers.getServerHost(service)

    def _getPort(self, service):
        return self._servers.getServerPort(service)

    def sendMessage(self, *args):
        Thread(target=self._sendMessage, args=args).start()

    def _sendMessage(self, recipient, subject, message, reset, progress, setstep):
        handler = RollerHandler(self._ctx, self._logger.Name)
        self._logger.addRollerHandler(handler)
        if self._user.hasThread():
            i = 0
            reset(100)
        elif self._hasImapService():
            i = 100
            reset(200)
            self._user.ThreadId = self._uploadMessage(subject, progress)
        smtp = SMTP.value
        context = self._getConnectionContext(smtp)
        authenticator = self._getAuthenticator(smtp)
        step = 3
        progress(i + 5)
        service = 'com.sun.star.mail.MailServiceProvider2'
        progress(i + 25)
        host = context.getValueByName('ServerName')
        server = createService(self._ctx, service).create(SMTP, host)
        progress(i + 50)
        try:
            server.connect(context, authenticator)
        except UnoException as e:
            print("IspdbModel._sendMessage() 1 Error: %s" % e.Message)
        else:
            progress(i + 75)
            if server.isConnected():
                body = MailTransferable(self._ctx, message, False)
                mail = getMail(self._ctx, self.Email, recipient, subject, body)
                if self._user.hasThread():
                    mail.ThreadId = self._user.ThreadId
                print("IspdbModel._sendMessage() 2: %s - %s" % (type(mail), mail))
                try:
                    server.sendMailMessage(mail)
                    print("IspdbModel._sendMessage() 3: %s" % mail.MessageId)
                except UnoException as e:
                    print("IspdbModel._sendMessage() 4 Error: %s - %s" % (e.Message, traceback.print_exc()))
                else:
                    step = 5
                server.disconnect()
        progress(i + 100)
        self._logger.removeRollerHandler(handler)
        setstep(step)

    def _uploadMessage(self, subject, progress):
        mail = msgid = None
        imap = IMAP.value
        context = self._getConnectionContext(imap)
        authenticator = self._getAuthenticator(imap)
        progress(10)
        service = 'com.sun.star.mail.MailServiceProvider2'
        host = context.getValueByName('ServerName')
        server = createService(self._ctx, service).create(IMAP, host)
        progress(20)
        try:
            server.connect(context, authenticator)
        except UnoException as e:
            print("IspdbModel._uploadMessage() 1 Error: %s" % e.Message)
        else:
            progress(40)
            if server.isConnected():
                try:
                    folder = server.getSentFolder()
                    if server.hasFolder(folder):
                        message = self._getThreadMessage()
                        body = MailTransferable(self._ctx, message, True)
                        mail = getMail(self._ctx, self.Email, self.Email, subject, body)
                        progress(60)
                        server.uploadMessage(folder, mail)
                except UnoException as e:
                    print("IspdbModel._uploadMessage() 2 Error: %s - %s" % (e.Message, traceback.print_exc()))
                else:
                    if mail is not None:
                        msgid = mail.MessageId
                progress(80)
                server.disconnect()
        progress(100)
        return msgid

# IspdbModel getter methods called by SendDialog
    def validSend(self, recipient, subject, message):
        enabled = all((self.isEmailValid(recipient),
                       self.isStringValid(subject),
                       self.isStringValid(message)))
        return enabled

# IspdbModel private getter methods called by WizardPage3 / WizardPage4
    def _getLogin(self, service):
        login = self._user.getLogin(service)
        return login if login else self._getLoginFromEmail(service)

    def _getLoginFromEmail(self, service):
        mode = self._servers.getLoginMode(service)
        return self.Email.partition('@')[mode] if mode != 1 else self.Email

    def _getPassword(self, service):
        return self._user.getPassword(service)

    def _isHostValid(self, host):
        if validators.domain(host):
            return True
        return False

    def _isPortValid(self, port):
        if validators.between(port, min=1, max=1023):
            return True
        return False

# IspdbModel private getter methods called by WizardPage5
    def _hasImapService(self):
        return IMAP.value in self._services

    def _getConnectionContext(self, service):
        config = self._servers.getConfig(service, self._timeout, self._connections, self._authentications)
        return CurrentContext(service, config)

    def _getAuthenticator(self, service):
        return Authenticator(service, self._user.getConfig())

    def _getThreadMessage(self):
        title = self._getThreadTitle()
        path = '%s/%s' % (g_extension, g_logo)
        url = getResourceLocation(self._ctx, g_identifier, path)
        logo = getMessageImage(self._ctx, url)
        return '''\
<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
  </head>
  <body>
    <img alt="%s Logo" src="data:image/png;charset=utf-8;base64,%s" src="%s" />
    <h3 style="display:inline;" >&nbsp;%s</h3>
  </body>
</html>
''' % (g_extension, logo, g_logourl, title)

# IspdbModel private shared methods
    def _getServicesCount(self):
        return len(self._services)

# IspdbModel StringResource methods
    def getPageStep(self, pageid):
        resource = self._resources.get('Step') % pageid
        return self._resolver.resolveString(resource)

    def getPageTitle(self, pageid):
        resource = self._resources.get('Title') % (pageid, self._offline)
        return self._resolver.resolveString(resource)

    def getPageLabel(self, pageid):
        resource = self._resources.get('PageLabel') % pageid
        return self._resolver.resolveString(resource)

    def getPagesLabel(self, service):
        resource = self._resources.get('PagesLabel')
        label = self._resolver.resolveString(resource)
        return label % (service, self.Email)

    def getProgressMessage(self, value):
        resource = self._resources.get('Progress') % value
        return self._resolver.resolveString(resource)

    def getSecurityMessage(self, level):
        resource = self._resources.get('Security') % level
        return self._resolver.resolveString(resource)

    def getSendTitle(self):
        resource = self._resources.get('SendTitle')
        title = self._resolver.resolveString(resource)
        return title % self.Email

    def getSendSubject(self):
        resource = self._resources.get('SendSubject')
        return self._resolver.resolveString(resource)

    def getSendMessage(self):
        resource = self._resources.get('SendMessage')
        return self._resolver.resolveString(resource)

    def _getThreadTitle(self):
        resource = self._resources.get('ThreadTitle')
        return self._resolver.resolveString(resource) % g_extension
