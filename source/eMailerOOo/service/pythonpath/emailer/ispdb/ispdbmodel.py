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

from com.sun.star.frame.DispatchResultState import FAILURE
from com.sun.star.frame.DispatchResultState import SUCCESS

from com.sun.star.mail.MailServiceType import IMAP
from com.sun.star.mail.MailServiceType import SMTP

from .task import Ispdb
from .task import Server
from .task import Send

from ..runner import Cancel
from ..runner import Runner

from ..unotool import findFrame
from ..unotool import getConfiguration
from ..unotool import getResourceLocation
from ..unotool import getStringResource

from ..helper import getMessageImage

from ..logger import LogController

from ..configuration import g_identifier
from ..configuration import g_extension
from ..configuration import g_img
from ..configuration import g_logo
from ..configuration import g_logourl
from ..configuration import g_mailservicelog
from ..configuration import g_sep
from ..configuration import g_spoolerframe

from email.utils import parseaddr
import validators


class IspdbModel():
    def __init__(self, ctx, source, sender, readonly, dispatch):
        self._ctx = ctx
        self._source = source
        self._sender = sender
        self._readonly = readonly
        self._dispatch = dispatch
        self._servers = None
        self._offline = 0
        self._diposed = False
        self._cancel = Cancel()
        self._version = 0
        self._check = False
        self._listener = None
        self._handler = False
        secure = {0: 3, 1: 4, 2: 5}
        unsecure = {0: 0, 1: 1, 2: 2}
        self._levels = {0: unsecure, 1: secure, 2: secure}
        self._timeout = getConfiguration(ctx, g_identifier).getByName('ConnectTimeout')
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
        _, address = parseaddr(self._sender)
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
        self._cancel.set(disposed=True)
        # XXX If notify is None the SUCCESS dispatch notification is already done
        if self._dispatch:
            struct = 'com.sun.star.beans.StringPair'
            result = uno.createUnoStruct(struct, 'Sender', '')
            struct = 'com.sun.star.frame.DispatchResultEvent'
            notification = uno.createUnoStruct(struct, self._source, FAILURE, result)
            self._dispatch.dispatchFinished(notification)
        if self._listener:
            self._logger.removeModifyListener(self._listener)
            self._listener = None
        if self._handler:
            self._logger.removeRollerHandler()

# IspdbModel getter methods called by WizardPages 1
    def getSender(self):
        return self._sender, not self._readonly

# IspdbModel getter methods called by WizardPages 2 and 4
    def isDisposed(self):
        return self._diposed

# IspdbModel getter methods called by WizardPage1
    def isEmailValid(self, sender):
        _, address = parseaddr(sender)
        if validators.email(address):
            return True
        return False

# IspdbModel getter methods called by WizardPage2
    def getServerConfig(self, callback, indicator):
        self._version += 1
        ispdb = Ispdb(self._ctx, callback, self._sender)
        runner = Runner(self._ctx, ispdb, self._cancel, indicator)
        runner.start()

# IspdbModel setter methods called by WizardPage2
    def setServerConfig(self, servers):
        self._servers = servers
        self._check = servers.hasConfig()

    def enableReplyTo(self, state):
        self._servers.User.enableReplyTo(bool(state))
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
        if self._dispatch:
            struct = 'com.sun.star.beans.StringPair'
            result = uno.createUnoStruct(struct, 'Sender', self._sender)
            struct = 'com.sun.star.frame.DispatchResultEvent'
            notification = uno.createUnoStruct(struct, self._source, SUCCESS, result)
            self._dispatch.dispatchFinished(notification)
            self._dispatch = None

    def getSecurity(self, resolver, i, j):
        level = self._levels.get(i).get(j)
        message = self.getSecurityMessage(resolver, level)
        return message, level

# IspdbModel getter methods called by WizardPage5
    def addLogListener(self, listener):
        self._logger.addModifyListener(listener)
        self._listener = listener
        # XXX: We add RollerHandler only if the Spooler is not loaded
        if findFrame(self._ctx, g_spoolerframe) is None:
            self._logger.addRollerHandler()
            self._handler = True

    def getMailServers(self):
        return self._servers

    def getMailServerLabel(self, resolver,  pageid, service):
        return self.getPageLabel(resolver,  pageid, service, self._getHost(service), self._getPort(service))

    def useIMAP(self):
        return self._servers.User.useIMAP()

    def getLogContent(self):
        return self._logger.getLogContent(True)

    def connectServer(self, callback, indicator, service):
        server = Server(self._ctx, callback, self._servers.User, service)
        runner = Runner(self._ctx, server, self._cancel, indicator)
        runner.start()

    def _getHost(self, service):
        return self._servers.getServerHost(service)

    def _getPort(self, service):
        return self._servers.getServerPort(service)

    def sendMail(self, callback, indicator, imap, recipient, subject, msg, thread=None):
        message = self._getThreadMessage() if imap else msg
        service = IMAP if imap else SMTP
        send = Send(self._ctx, callback, self._servers.User, service, self._sender, recipient, subject, message, thread)
        runner = Runner(self._ctx, send, self._cancel, indicator)
        runner.start()

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
        valid = False
        if self._check and validators.domain(host):
            valid = True
        elif not self._check and validators.hostname(host):
            valid = True
        return valid

    def _isPortValid(self, port):
        valid = False
        if self._check and validators.between(port, min_val=0, max_val=1023):
            valid = True
        elif not self._check and validators.between(port, min_val=0, max_val=65535):
            valid = True
        return valid

# IspdbModel private getter methods called by WizardPage5
    def _getThreadMessage(self):
        title = self._getThreadTitle()
        url = getResourceLocation(self._ctx, g_identifier, g_img + g_sep + g_logo)
        logo = getMessageImage(self._ctx, url)
        return '''\
<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
  </head>
  <body>
    <table border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 15px; border-collapse: collapse;">
      <tr>
        <td valign="top" width="64" style="padding-right: 10px; width: 64px;">
          <img alt="%s Logo" src="data:image/png;charset=utf-8;base64,%s" src="%s" width="64" height="64" style="display: block; width: 64px; height: 64px; border: 0;" />
        </td>
        <td valign="top" height="64" style="height: 64px;">
          <table border="0" cellpadding="0" cellspacing="0" height="64" style="height: 64px; border-collapse: collapse;">
            <tr>
              <td valign="top" style="line-height: 1.2;">
                <span style="font-size: 1.5em; font-weight: bold; font-family: sans-serif; display: block; color: #000000; line-height: 1.2;">%s</span>
              </td>
            </tr>
            <tr>
              <td valign="bottom" style="line-height: 1.2;">
                <span style="font-size: 1.15em; font-weight: bold; font-family: sans-serif; display: block; color: #333333; line-height: 1.2;">%s</span>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>
''' % (g_extension, logo, g_logourl, g_extension, title)

# IspdbModel StringResource methods
    def getPageStep(self, resolver, pageid):
        resource = self._resources.get('Step') % pageid
        return resolver.resolveString(resource)

    def getPageTitle(self, resolver, pageid):
        resource = self._resources.get('Title') % (pageid, self._offline)
        return resolver.resolveString(resource)

    def getPageLabel(self, resolver, pageid, *options):
        resource = self._resources.get('PageLabel') % pageid
        return resolver.resolveString(resource) % options

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
        return self._resolver.resolveString(resource)

    def _getResolver(self, name):
        return getStringResource(self._ctx, g_identifier, 'dialogs', name)
