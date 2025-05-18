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

from ..transferable import Transferable

from ..unotool import createService
from ..unotool import getConfiguration
from ..unotool import getConnectionMode
from ..unotool import getResourceLocation
from ..unotool import getStringResource
from ..unotool import getUrl
from ..unotool import hasInterface

from ..helper import getMailService
from ..helper import getMailMessage
from ..helper import getMessageImage

from ..logger import LogController
from ..logger import RollerHandler

from ..oauth20 import g_service

from ..configuration import g_identifier
from ..configuration import g_extension
from ..configuration import g_mailservicelog
from ..configuration import g_logo
from ..configuration import g_logourl
from ..configuration import g_chunk

from ..user import User

from .ispdbserver import IspdbServer

from email.utils import parseaddr
import xml.etree.ElementTree as ET
from time import sleep
from threading import Thread
import validators
import traceback


class IspdbModel(unohelper.Base):
    def __init__(self, ctx, sender, readonly):
        self._ctx = ctx
        self._sender = sender
        self._readonly = readonly
        self._servers = None
        self._services = [SMTP.value, IMAP.value]
        self._offline = 0
        self._diposed = False
        self._version = 0
        self._listener = None
        secure = {0: 3, 1: 4, 2: 5}
        unsecure = {0: 0, 1: 1, 2: 2}
        self._levels = {0: unsecure, 1: secure, 2: secure}
        self._connections = {0: 'Insecure', 1: 'SSL', 2: 'TLS'}
        self._authentications = {0: 'None', 1: 'Login', 2: 'OAuth2'}
        config = getConfiguration(ctx, g_identifier)
        self._url = config.getByName('IspDBUrl')
        self._timeout = config.getByName('ConnectTimeout')
        self._links = self._getAppsLinks(config.getByName('Applications'))
        self._clients = config.getByName('Urls').getElementNames()
        self._providers = config.getByName('Providers')
        self._logger = LogController(ctx, g_mailservicelog)
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'MessageBox')
        self._resources = {'Step':        'IspdbPage%s.Step',
                           'Title':       'IspdbPage%s.Title.%s',
                           'PageLabel':   'IspdbPage%s.Label1.Label',
                           'PagesLabel':  'IspdbPages.Label1.Label',
                           'Progress':    'IspdbPage2.Label2.Label.%s',
                           'Security':    'IspdbPages.Label10.Label.%s',
                           'SendTitle':   'SendDialog.Title',
                           'SendSubject': 'SendDialog.TextField2.Text',
                           'SendMessage': 'SendDialog.TextField3.Text',
                           'ThreadTitle': 'SendThread.Title'}

    @property
    def Email(self):
        name, address = parseaddr(self._sender)
        return address

    @property
    def Sender(self):
        return self._sender
    @Sender.setter
    def Sender(self, sender):
        self._sender = sender

    @property
    def Offline(self):
        return self._offline
    @Offline.setter
    def Offline(self, offline):
        self._offline = offline

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

# IspdbModel getter methods called by IspdbWizard
    def dispose(self):
        self._diposed = True
        if self._listener is not None:
            self._logger.removeModifyListener(self._listener)
            self._listener = None

# IspdbModel getter methods called by WizardPages 1
    def getSender(self):
        return self._sender, not self._readonly

# IspdbModel getter methods called by WizardPages 2 and 4
    def isDisposed(self):
        return self._diposed

# IspdbModel getter methods called by WizardPage1
    def isEmailValid(self, sender):
        name, address = parseaddr(sender)
        if validators.email(address):
            return True
        return False

# IspdbModel getter methods called by WizardPage2
    def getServerConfig(self, *args):
        self._version += 1
        Thread(target=self._getServerConfig, args=args).start()

    def _getServerConfig(self, resolver, progress, update):
        # FIXME: Because we call this thread in the WizardPage.activatePage(),
        # FIXME: if we want to be able to navigate through the Wizard roadmap
        # FIXME: without GUI refreshing problem then we need to pause this thread.
        sleep(0.2)
        auto = False
        config = None
        progress(10)
        url = getUrl(self._ctx, self._url)
        progress(20)
        mode = getConnectionMode(self._ctx, url.Server)
        progress(30)
        user = User(self._ctx, self.Sender)
        progress(40)
        request = createService(self._ctx, g_service)
        progress(50)
        if request is None:
            offset = 1 if user.isNew() else 2
            status = (100, offset, BOLD)
        elif mode == OFFLINE:
            offset = 3 if user.isNew() else 4
            status = (100, offset, BOLD)
        else:
            progress(60)
            try:
                config = self._getIspdbConfig(user, request, url.Complete)
            except RequestException as e:
                offset = 5 if user.isNew() else 6
                status = (100, offset, BOLD)
            else:
                progress(80)
                if not user.isNew():
                    status = (100, 7)
                elif config is None:
                    status = (100, 8)
                else:
                    auto = True
                    status = (100, 9)
        self._servers = IspdbServer(user, config)
        self._offline = mode
        progress(*status)
        update(self.getPageTitle(resolver, 2), auto, user.getReplyToState(), user.ReplyToAddress, user.getImapState(), mode)

    def _getIspdbConfig(self, user, request, url):
        config = None
        parameter = request.getRequestParameter('getIspdbConfig')
        parameter.Url = '%s%s' % (url, user.getUserDomain())
        parameter.NoAuth = True
        response = request.execute(parameter)
        if response.Ok:
            config = self._parseIspdbConfig(user, request, response)
        elif response.StatusCode != NOT_FOUND:
            response.raiseForStatus()
        response.close()
        return config

    def _parseIspdbConfig(self, user, request, response):
        smtps = []
        imaps = []
        config = {}
        provider = None
        map1 = {'none': 0, 'password-cleartext': 1, 'plain': 1,
                'password-encrypted': 1, 'secure': 1, 'OAuth2': 2}
        map2 = {'plain': 0, 'SSL': 1, 'STARTTLS': 2}
        map3 = {'%EMAILLOCALPART%': 0, '%EMAILADDRESS%': 1, '%EMAILDOMAIN%': 2}
        parser = ET.XMLPullParser(('end', ))
        iterator = response.iterContent(g_chunk, False)
        while iterator.hasMoreElements():
            # FIXME: As Decode is False we obtain a sequence of bytes
            parser.feed(iterator.nextElement().value)
            for event, element in parser.read_events():
                if element.tag == 'emailProvider':
                    provider = element.get('id')
                    for child in element.findall('outgoingServer'):
                        if child.get('type') == 'smtp':
                            smtps.append(self._parseServer(request, provider, child, map1, map2, map3))
                    for child in element.findall('incomingServer'):
                        if child.get('type') == 'imap':
                            imaps.append(self._parseServer(request, provider, child, map1, map2, map3))
        config[SMTP.value] = tuple(smtps)
        config[IMAP.value] = tuple(imaps)
        if user.isNew():
            user.enableImap(len(imaps) > 0)
            if provider in self._clients:
                user.setClient(provider)
        # XXX: If provider use API and we need to be able to test the connection
        # XXX: then it is necessary to save the domain used by the user
        if provider:
            user.setProvider(provider)
            user.setAppsLink(self._links.get(provider, ''))
        return config

    def _parseServer(self, request, provider, element, map1, map2, map3):
        hostname = element.find('hostname').text
        authentication = map1.get(element.find('authentication').text, 1)
        if authentication == 2 and not self._supportOAuth2(request, provider, hostname):
            authentication = 1
        server = {'ServerName':         hostname,
                  'Port':               int(element.find('port').text),
                  'AuthenticationType': authentication,
                  'ConnectionType':     map2.get(element.find('socketType').text, 1),
                  'LoginMode':          map3.get(element.find('username').text, 1)}
        return server

    def _supportOAuth2(self, request, provider, hostname):
        # FIXME: We support OAuth2 protocol only for providers:
        # FIXME: - Who have registered their hostname as a ResourceURL
        # FIXME: - Which uses an API to replace SMTP and/or IMAP servers
        return self._providers.hasByName(provider) or request.isRegisteredUrl(hostname)

# IspdbModel setter methods called by WizardPage2
    def enableReplyTo(self, enabled):
        self._servers.User.enableReplyTo(enabled)
        return self._servers.User.ReplyToAddress

    def setReplyToAddress(self, replyto):
        self._servers.User.setReplyToAddress(replyto)

    def enableIMAP(self, imap):
        self._servers.User.enableImap(bool(imap))
        return self._offline

# IspdbModel getter methods called by WizardPage3 / WizardPage4
    def refreshView(self, version):
        return version < self._version

    def getVersion(self):
        return self._version

    def getConfig(self, service):
        config = {}
        config['First'] = self._servers.isFirst(service)
        config['Last'] = self._servers.isLast(service)
        config['Page'] = self._servers.getServerPage(service)
        config['Default'] = self._servers.isDefaultPage(service)
        config.update(self._servers.getCurrentServer(service))
        config['UserName'] = self._getUserName(service)
        config['Password'] = self._getPassword(service)
        config['AppsLink'] = self._servers.User.getAppsLink()
        return config

    def isConnectionValid(self, host, port):
        return self._isHostValid(host) and self._isPortValid(port)

    def isStringValid(self, value):
        # FIXME: validators 0.22.0 seem to be buggy, as workaround we need to set max_val
        # FIXME: see: https://github.com/python-validators/validators/issues/323)
        max_val = max(1, len(value))
        if validators.length(value, min_val=1, max_val=max_val):
            return True
        return False

    def previousServerPage(self, service, server):
        self._servers.updateCurrentServer(service, server)
        self._servers.decrementIndex(service)

    def nextServerPage(self, service, server):
        self._servers.updateCurrentServer(service, server)
        self._servers.incrementIndex(service)

    def updateConfiguration(self, service, server, user):
        self._servers.updateCurrentServer(service, server)
        self._servers.User.updateConfig(service, user)

    def saveConfiguration(self):
        self._servers.User.saveConfig()

    def getSecurity(self, resolver, i, j):
        level = self._levels.get(i).get(j)
        message = self.getSecurityMessage(resolver, level)
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
        level = self._getLogLevel(handler)
        i = 0
        connect = True
        range = 100
        count = 2 if self._servers.User.useIMAP() else 1
        reset(count * range)
        for service in self._services:
            if service == SMTP.value or self._servers.User.useIMAP():
                setlabel(service, self._getHost(service), self._getPort(service))
                context = self._servers.User.getConnectionContext(service)
                authenticator = self._servers.User.getAuthenticator(service)
                if not self._connectServer(context, authenticator, service, i * range, progress):
                    connect = False
                    break
                i += 1
        self._setLogLevel(handler, level)
        setstep(4 if connect else 2)

    def _connectServer(self, context, authenticator, service, i, progress):
        connect = False
        progress(i + 25)
        server = getMailService(self._ctx, service)
        progress(i + 50)
        try:
            server.connect(context, authenticator)
        except UnoException as e:
            # Exceptions are already logged in the MailServiceLogger log.
            progress(i + 100)
        else:
            progress(i + 75)
            if server.isConnected():
                server.disconnect()
                connect = True
                progress(i + 100)
            else:
                progress(i + 100)
        return connect

    def _getHost(self, service):
        return self._servers.getServerHost(service)

    def _getPort(self, service):
        return self._servers.getServerPort(service)

    def sendMessage(self, *args):
        Thread(target=self._sendMessage, args=args).start()

    def _sendMessage(self, recipient, subject, message, reset, progress, setstep):
        handler = RollerHandler(self._ctx, self._logger.Name)
        level = self._getLogLevel(handler)
        transferable = Transferable(self._ctx, self._logger)
        if self._servers.User.useIMAP():
            i = 100
            reset(200)
            mail = self._uploadMessage(transferable, subject, progress)
            threadid = mail.ForeignId if mail.ForeignId else mail.MessageId
        else:
            i = 0
            reset(100)
            threadid = None
        progress(i + 5)
        smtp = SMTP.value
        context = self._servers.User.getConnectionContext(smtp)
        authenticator = self._servers.User.getAuthenticator(smtp)
        step = 3
        progress(i + 25)
        server = getMailService(self._ctx, smtp)
        progress(i + 50)
        try:
            server.connect(context, authenticator)
        except UnoException as e:
            # Exceptions are already logged in the MailServiceLogger log.
            print("IspdbModel._sendMessage() Error: %s - %s" % (e.Message, traceback.format_exc()))
        else:
            progress(i + 75)
            if server.isConnected():
                body = transferable.getByString(message)
                mail = getMailMessage(self._ctx, self.Sender, recipient, subject, body)
                interface = 'com.sun.star.mail.XMailMessage2'
                # If IMAP is supported then we set the Message.ThreadId
                if threadid and hasInterface(mail, interface):
                    mail.ThreadId = threadid
                try:
                    server.sendMailMessage(mail)
                except UnoException as e:
                    # Exceptions are already logged in the MailServiceLogger log.
                    print("IspdbModel._sendMessage() Error: %s - %s" % (e.Message, traceback.format_exc()))
                else:
                    step = 5
                server.disconnect()
        progress(i + 100)
        self._setLogLevel(handler, level)
        setstep(step)

    def _uploadMessage(self, transferable, subject, progress):
        mail = None
        imap = IMAP.value
        context = self._servers.User.getConnectionContext(imap)
        authenticator = self._servers.User.getAuthenticator(imap)
        progress(10)
        server = getMailService(self._ctx, imap)
        progress(20)
        try:
            server.connect(context, authenticator)
        except UnoException as e:
            # Exceptions are already logged in the MailServiceLogger log.
            print("IspdbModel._uploadMessage() Error: %s - %s" % (e.Message, traceback.format_exc()))
        else:
            progress(40)
            if server.isConnected():
                try:
                    folder = server.getSentFolder()
                    if server.hasFolder(folder):
                        message = self._getThreadMessage()
                        body = transferable.getByString(message)
                        mail = getMailMessage(self._ctx, self.Sender, self.Email, subject, body)
                        progress(60)
                        server.uploadMessage(folder, mail)
                except UnoException as e:
                    # Exceptions are already logged in the MailServiceLogger log.
                    print("IspdbModel._uploadMessage() Error: %s - %s" % (e.Message, traceback.format_exc()))
                progress(80)
                server.disconnect()
        progress(100)
        return mail

    def _getLogLevel(self, handler):
        self._logger.addRollerHandler(handler)
        level = self._logger.Level
        self._logger.Level = ALL
        return level

    def _setLogLevel(self, handler, level):
        self._logger.removeRollerHandler(handler)
        self._logger.Level = level

# IspdbModel getter methods called by SendDialog
    def getSendMailData(self):
        title = self._getSendTitle('SendDialog')
        subject = self._getSendSubject('SendDialog')
        msg = self._getSendMessage('SendDialog')
        return self.Email, title, subject, msg

    def validSend(self, recipient, subject, message):
        enabled = all((self.isEmailValid(recipient),
                       self.isStringValid(subject),
                       self.isStringValid(message)))
        return enabled

# IspdbModel private getter methods called by WizardPage3 / WizardPage4
    def _getUserName(self, service):
        if self._servers.User.isNew():
            username = self._getUserNameFromEmail(service)
        else:
            username = self._servers.User.getUserName(service)
        return username

    def _getUserNameFromEmail(self, service):
        mode = self._servers.getLoginMode(service)
        return self.Email.partition('@')[mode] if mode != 1 else self.Email

    def _getPassword(self, service):
        return self._servers.User.getPassword(service)

    def _isHostValid(self, host):
        if validators.domain(host):
            return True
        return False

    def _isPortValid(self, port):
        if validators.between(port, min_val=1, max_val=1023):
            return True
        return False

# IspdbModel private getter methods called by WizardPage5
    def _getThreadMessage(self):
        title = self._getThreadTitle()
        path = '%s/%s' % (g_extension, g_logo)
        url = getResourceLocation(self._ctx, g_identifier, path)
        logo = getMessageImage(self._ctx, url)
        msg = '''\
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
</html>''' % (g_extension, logo, g_logourl, title)
        return msg

# IspdbModel private shared methods
    def _getServicesCount(self):
        return len(self._services)

    def _getAppsLinks(self, applications):
        links = {}
        for name in applications.getElementNames():
            apps = applications.getByName(name)
            links[name] = apps.getByName('Link')
        return links

# IspdbModel StringResource methods
    def getPageStep(self, resolver, pageid):
        resource = self._resources.get('Step') % pageid
        return resolver.resolveString(resource)

    def getPageTitle(self, resolver, pageid):
        resource = self._resources.get('Title') % (pageid, self._offline)
        return resolver.resolveString(resource)

    def getPageLabel(self, resolver, pageid, *format):
        resource = self._resources.get('PageLabel') % pageid
        return resolver.resolveString(resource) % format

    def getPagesLabel(self, resolver, service):
        resource = self._resources.get('PagesLabel')
        label = resolver.resolveString(resource)
        return label % (service, self.Email)

    def getProgressMessage(self, resolver, value):
        resource = self._resources.get('Progress') % value
        return resolver.resolveString(resource)

    def getSecurityMessage(self, resolver, level):
        resource = self._resources.get('Security') % level
        return resolver.resolveString(resource)

    def _getSendTitle(self, name):
        resource = self._resources.get('SendTitle')
        title = self._getResolver(name).resolveString(resource)
        return title % self.Email

    def _getSendSubject(self, name):
        resource = self._resources.get('SendSubject')
        return self._getResolver(name).resolveString(resource)

    def _getSendMessage(self, name):
        resource = self._resources.get('SendMessage')
        return self._getResolver(name).resolveString(resource)

    def _getThreadTitle(self):
        resource = self._resources.get('ThreadTitle')
        return self._resolver.resolveString(resource) % g_extension

    def _getResolver(self, name):
        return getStringResource(self._ctx, g_identifier, 'dialogs', name)
