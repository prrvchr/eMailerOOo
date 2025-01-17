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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .cardtool import getSqlException

from .oauth20 import getRequest

from dateutil import parser
from dateutil import tz

import traceback

class Provider():
    def __init__(self, ctx):
        self._cls = 'Provider'
        self._ctx = ctx

    # Currently only vCardOOo supports multiple address books
    def supportAddressBook(self):
        return False

    # Currently only vCardOOo does not supports group
    def supportGroup(self):
        return True

    def parseDateTime(self, timestamp):
        datetime = uno.createUnoStruct('com.sun.star.util.DateTime')
        try:
            dt = parser.parse(timestamp)
        except parser.ParserError:
            pass
        else:
            datetime.Year = dt.year
            datetime.Month = dt.month
            datetime.Day = dt.day
            datetime.Hours = dt.hour
            datetime.Minutes = dt.minute
            datetime.Seconds = dt.second
            datetime.NanoSeconds = dt.microsecond * 1000
            datetime.IsUTC = dt.tzinfo == tz.tzutc()
        return datetime

    # Method called from User.__init__()
    # This main method call Request with OAuth2 mode
    def getRequest(self, url, name):
        return getRequest(self._ctx, url, name)

    # Need to be implemented method
    def insertUser(self, source, database, request, scheme, server, name, pwd):
        raise NotImplementedError

    # Method called from DataSource.getConnection()
    def initAddressbooks(self, source, logger, database, user):
        raise NotImplementedError

    def initUserBooks(self, source, logger, database, user, books):
        count = 0
        modified = False
        mtd = 'initUserBooks'
        logger.logprb(INFO, self._cls, mtd, 1331, user.Name)
        for uri, name, tag, token in books:
            print("Provider.initUserBooks() 1 Name: %s - Uri: %s - Tag: %s - Token: %s" % (name, uri, tag, token))
            if user.hasBook(uri):
                book = user.getBook(uri)
                if book.hasNameChanged(name):
                    database.updateAddressbookName(book.Id, name)
                    book.setName(name)
                    modified = True
                    print("Provider.initUserBooks() 2 %s" % (name, ))
            else:
                args = database.insertBook(user.Id, uri, name, tag, token)
                if args:
                    user.setNewBook(uri, **args)
                    modified = True
                    print("Provider.initUserBooks() 3 %s - %s - %s" % (user.getBook(uri).Id, name, uri))
            self.initUserGroups(source, logger, database, user, uri)
            count += 1
        print("Provider.initUserBooks() 4")
        if not count:
            raise getSqlException(self._ctx, source, 1006, 1611, self._cls, 'initUserBooks', user.Name, user.Server)
        if modified and self.supportAddressBook():
            database.initAddressbooks(user)
        logger.logprb(INFO, self._cls, mtd, 1332, user.Name)

    def initUserGroups(self, source, logger, database, user, uri):
        raise NotImplementedError

    def firstPullCard(self, database, user, addressbook, pages, count):
        raise NotImplementedError

    def pullCard(self, database, user, addressbook, pages, count):
        raise NotImplementedError

    def parseCard(self, database):
        raise NotImplementedError

    def raiseForStatus(self, source, mtd, response, user):
        name = response.Parameter.Name
        url = response.Parameter.Url
        status = response.StatusCode
        msg = response.Text
        response.close()
        raise getSqlException(self._ctx, source, 1006, 1601, self._cls, mtd,
                              name, status, user, url, msg)

    def getLoggerArgs(self, response, mtd, parameter, user):
        status = response.StatusCode
        msg = response.Text
        response.close()
        return ['Provider', mtd, 201, parameter.Name, status, user, parameter.Url, msg]

    # Can be overwritten method
    def syncGroups(self, database, user, addressbook, pages, count):
        return pages, count, None

