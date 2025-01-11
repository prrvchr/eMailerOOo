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

from com.sun.star.rest.ParameterType import REDIRECT

from ..provider import Provider as ProviderMain

from ..dbtool import currentDateTimeInTZ

from ..configuration import g_host
from ..configuration import g_url
from ..configuration import g_chunk
from ..configuration import g_userfields
from ..configuration import g_groupfields

import json
import ijson
import traceback


class Provider(ProviderMain):
    def __init__(self, paths, lists, maps, types, tmps, fields):
        self._paths = paths
        self._lists = lists
        self._maps = maps
        self._types = types
        self._tmps = tmps
        self._fields = fields

    @property
    def Host(self):
        return g_host
    @property
    def BaseUrl(self):
        return g_url

# Method called from DataSource.getConnection()
    def getUserUri(self, name):
        return name

    def initAddressbooks(self, source, database, user):
        parameter = self._getRequestParameter(user.Request, 'getBooks')
        response = user.Request.execute(parameter)
        if not response.Ok:
            cls, mtd = 'Provider', 'initAddressbooks()'
            self.raiseForStatus(source, cls, mtd, response, user.Name)
        iterator = self._parseAllBooks(response)
        self.initUserBooks(source, database, user, iterator)

    def initUserGroups(self, database, user, book):
        parameter = self._getRequestParameter(user.Request, 'getGroups')
        response = user.Request.execute(parameter)
        if not response.Ok:
            cls, mtd = 'Provider', 'initUserGroups()'
            self.raiseForStatus(source, cls, mtd, response, user.Name)
        iterator = self._parseGroups(response)
        remove, add = database.initGroups(book, iterator)
        database.initGroupView(user, remove, add)

    # Method called from User.__init__()
    def insertUser(self, source, database, request, name, pwd):
        userid = self._getNewUserId(source, request, name, pwd)
        return database.insertUser(userid, '', name)
 
    # Private method
    def _getNewUserId(self, source, request, name, pwd):
        parameter = self._getRequestParameter(request, 'getUser')
        response = request.execute(parameter)
        if not response.Ok:
            cls, mtd = 'Provider', '_getNewUserId()'
            self.raiseForStatus(source, cls, mtd, response, name)
        userid = self._parseUser(response)
        return userid

    def _parseUser(self, response):
        userid = None
        events = ijson.sendable_list()
        parser = ijson.parse_coro(events)
        iterator = response.iterContent(g_chunk, False)
        while iterator.hasMoreElements():
            parser.send(iterator.nextElement().value)
            for prefix, event, value in events:
                if (prefix, event) == ('id', 'string'):
                    userid = value
            del events[:]
        parser.close()
        response.close()
        return userid

    def _parseAllBooks(self, response):
        events = ijson.sendable_list()
        parser = ijson.parse_coro(events)
        url = name = tag = None
        iterator = response.iterContent(g_chunk, False)
        while iterator.hasMoreElements():
            parser.send(iterator.nextElement().value)
            for prefix, event, value in events:
                if (prefix, event) == ('value.item.id', 'string'):
                    url = value
                elif (prefix, event) == ('value.item.displayName', 'string'):
                    name = value
                elif (prefix, event) == ('value.item.parentFolderId', 'string'):
                    tag = value
                if all((url, name, tag)):
                    yield  url, name, tag, ''
                    url = name = tag = None
            del events[:]
        parser.close()
        response.close()



# Method called from Replicator.run()
    def firstPullCard(self, database, user, book, page, count):
        return self._pullCard(database, 'firstPullCard()', user, book, page, count)

    def pullCard(self, database, user, book, page, count):
        return self._pullCard(database, 'pullCard()', user, book, page, count)

    def parseCard(self, database):
        start = database.getLastUserSync()
        stop = currentDateTimeInTZ()
        iterator = self._parseCardValue(database, start, stop)
        count = database.mergeCardValue(iterator)
        database.updateUserSync(stop)

    # Private method
    def _pullCard(self, database, mtd, user, book, page, count):
        args = []
        parameter = self._getRequestParameter(user.Request, 'getCards', book.Uri)
        iterator = self._parseCards(user, parameter, mtd, args)
        count += database.mergeCard(book.Id, iterator)
        page += parameter.PageCount
        if not args:
            if parameter.SyncToken:
                database.updateAddressbookToken(book.Id, parameter.SyncToken)
        #self.initGroups(database, user, book)
        return page, count, args

    def _parseCards(self, user, parameter, mtd, args):
        map = tmp = False
        while parameter.hasNextPage():
            response = user.Request.execute(parameter)
            if not response.Ok:
                args += self.getLoggerArgs(response, mtd, parameter, user.Name)
                break
            events = ijson.sendable_list()
            parser = ijson.parse_coro(events)
            url = tag = data = None
            iterator = response.iterContent(g_chunk, False)
            while iterator.hasMoreElements():
                parser.send(iterator.nextElement().value)
                for prefix, event, value in events:
                    if (prefix, event) == ('@odata.nextLink', 'string'):
                        parameter.setNextPage('', value, REDIRECT)
                    elif (prefix, event) == ('@odata.deltaLink', 'string'):
                        parameter.SyncToken = value
                    elif (prefix, event) == ('value.item', 'start_map'):
                        cid = etag = tmp = label = None
                        data = {}
                        deleted = False
                    elif (prefix, event) == ('value.item.deleted', 'boolean'):
                        deleted = value
                    elif (prefix, event) == ('value.item.id', 'string'):
                        cid = value
                    elif (prefix, event) == ('value.item.@odata.etag', 'string'):
                        etag = value
                    # FIXME: All the data parsing is done based on the tables: Resources, Properties and Types 
                    # FIXME: Only properties listed in these tables will be parsed
                    # FIXME: This is the part for simple property import (use of tables: Resources and Properties)
                    elif event == 'string' and prefix in self._paths:
                        data[self._paths.get(prefix)] = value
                    # FIXME: This is the part for simple list property import (use of tables: Resources and Properties)
                    elif event == 'start_array' and value is None and prefix + '.item' in self._lists:
                        data[self._lists.get(prefix + '.item')] = []
                    elif event == 'string' and prefix in self._lists:
                        data[self._lists.get(prefix)].append(value)
                    # FIXME: This is the part for typed property import (use of tables: Resources, Properties and Types)
                    elif event == 'start_map' and prefix in self._maps:
                        map = tmp = None
                        suffix = ''
                    elif event == 'map_key' and prefix in self._maps and value in self._maps.get(prefix):
                        suffix = value
                    elif event == 'string' and map is None and prefix in self._types:
                        map = self._types.get(prefix).get(value + suffix)
                    elif event == 'string' and tmp is None and prefix in self._tmps:
                        tmp = value
                    elif event == 'end_map' and map and tmp and prefix in self._maps:
                        data[map] = tmp
                        map = tmp = False
                    elif (prefix, event) == ('value.item', 'end_map'):
                        yield cid, etag, deleted, json.dumps(data)
                del events[:]
            parser.close()
            response.close()

    def _pullGroup(self, database, mtd, user, addressbook, page, count):
        parameter = self._getRequestParameter(user.Request, 'getGroups', addressbook)
        response = user.Request.execute(parameter)
        if not response.Ok:
            args = self.getLoggerArgs(response, mtd, parameter, user.Name)
            return page, count, args
        iterator = self._parseGroups(response)
        count += database.mergeGroup(addressbook.Id, iterator)
        page += parameter.PageCount
        return page, count, []

    def _parseGroups(self, response):
        events = ijson.sendable_list()
        parser = ijson.parse_coro(events)
        iterator = response.iterContent(g_chunk, False)
        while iterator.hasMoreElements():
            parser.send(iterator.nextElement().value)
            for prefix, event, value in events:
                if (prefix, event) == ('value.item', 'start_map'):
                    uri = name = None
                elif (prefix, event) == ('value.item.id', 'string'):
                    uri = value
                elif (prefix, event) == ('value.item.displayName', 'string'):
                    name = value
                elif (prefix, event) == ('value.item', 'end_map'):
                    if all((uri, name)):
                        yield  uri, name
            del events[:]
        parser.close()
        response.close()

    def _parseCardValue(self, database, start, stop):
        indexes = database.getColumnIndexes({'categories': -1})
        for book, card, query, data in database.getChangedCard(start, stop):
            if query == 'Deleted':
                continue
            else:
                for column, value in json.loads(data).items():
                    yield book, card, indexes.get(column), value



    def _getRequestParameter(self, request, method, data=None):
        parameter = request.getRequestParameter(method)
        parameter.Url = self.BaseUrl

        if method == 'getUser':
            parameter.Url += '/me'
            parameter.setQuery('select', g_userfields)

        elif method == 'getBooks':
            parameter.Url += '/me/contactFolders'

        elif method == 'getCards':
            parameter.Url += '/me/contactFolders/%s/contacts' % data
            parameter.setQuery('select', self._fields)

        elif method == 'getGroups':
            parameter.Url += '/me/outlook/masterCategories'
            parameter.setQuery('select', g_groupfields)

        elif method == 'getModifiedCardByToken':
            parameter.Url += '/me/contactFolders/%s/contacts/delta' % data
            parameter.setQuery('select', self._fields)

        return parameter

