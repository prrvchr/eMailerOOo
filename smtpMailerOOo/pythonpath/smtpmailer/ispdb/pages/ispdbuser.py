#!
# -*- coding: utf_8 -*-

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

from com.sun.star.mail.MailServiceType import SMTP
from com.sun.star.mail.MailServiceType import IMAP

import traceback


class IspdbUser(unohelper.Base):
    def __init__(self, user):
        self._user = user
        self._metadata = user.toJson()
        self._smtp = SMTP.value
        self._imap = IMAP.value

    @property
    def ThreadId(self):
        return self._user.getValue('ThreadId')
    @ThreadId.setter
    def ThreadId(self, thread):
        self._user.setValue('ThreadId', thread)

    def updateUser(self, user):
        self._user.update(user)

    def isUpdated(self):
        return self._user.toJson() != self._metadata

    def getServer(self, service):
        return self._user.getValue(service + 'Server')

    def getPort(self, service):
        return self._user.getValue(service + 'Port')

    def getLogin(self, service):
        return self._user.getValue(service + 'Login')

    def getPassword(self, service):
        return self._user.getValue(service + 'Password')

    def hasThread(self):
        return self.ThreadId is not None

    def hasImapConfig(self):
        return all((self.getServer(self._imap),
                    self.getPort(self._imap),
                    self.getLogin(self._imap)))

    def getDomain(self):
        return self._user.getValue('Domain')

    def getConfig(self):
        return self._user
