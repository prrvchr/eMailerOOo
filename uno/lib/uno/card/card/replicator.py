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

from com.sun.star.uno import Exception as UnoException

from ..runner import Task

from ..unotool import getNamedValue


class Replicator(Task):
    def __init__(self, ctx, /, database, provider, users, job, listener, timeout):
        super().__init__(ctx, 'Replicator')
        self._database = database
        self._provider = provider
        self._users = users
        self._job = job
        self._listener = listener
        self._timeout = timeout

        self.status = 'SUCCESS'
        self.error_cause = None

        self.stats_card = {'users': 0, 'pages': 0, 'count': 0}
        self.stats_group = {'users': 0, 'pages': 0, 'count': 0}

    @property
    def steps(self):
        if not self._hasConnectedUser():
            self.status = 'NO_USER'
            return

        # --- PHASE 1 : SYNCHRONISATION DES CARTES ---
        for user in self._users.values():
            if user.isOffLine():
                continue
            for book in user.getBooks():
                if book.isNew():
                    yield (11103, (user.Name,), self._execFirstPull, user, book)
                else:
                    yield (11103, (user.Name,), self._execPull, user, book)

        yield (11101, (), self._execParseCard)

        # --- PHASE 2 : SYNCHRONISATION DES GROUPES ---
        if self._provider.supportGroup():
            for user in self._users.values():
                if user.isOffLine():
                    continue
                for book in user.getBooks():
                    yield (12103, (user.Name,), self._execSyncGroup, user, book)
                    yield (12104, (user.Name,), self._execDbSyncGroup, user)

    def _execFirstPull(self, user, book):
        self.stats_card['users'] += 1
        p, c, args = self._provider.firstPullCard(self._database, user, book, self.stats_card['pages'], self.stats_card['count'])
        if args:
            raise UnoException(f'Erreur FirstPull : {args}')
        self.stats_card['pages'] = p
        self.stats_card['count'] = c
        book.resetNew()

    def _execPull(self, user, book):
        self.stats_card['users'] += 1
        p, c, args = self._provider.pullCard(self._database, user, book, self.stats_card['pages'], self.stats_card['count'])
        if args:
            raise UnoException(f'Erreur Pull : {args}')
        self.stats_card['pages'] = p
        self.stats_card['count'] = c
        book.resetNew()

    def _execParseCard(self):
        if self.stats_card['count'] > 0:
            self._provider.parseCard(self._database)

    def _execSyncGroup(self, user, book):
        self.stats_group['users'] += 1
        p, c, args = self._provider.syncGroups(self._database, user, book, self.stats_group['pages'], self.stats_group['count'])
        if args:
            raise UnoException(f'Erreur SyncGroup : {args}')
        self.stats_group['pages'] = p
        self.stats_group['count'] = c

    def _execDbSyncGroup(self, user):
        self._database.syncGroups(user)

    def getStepFinalize(self):
        return (10104, (), self._database.dispose)

    def isExtended(self, total):
        if self.total == 0:
            total_steps = 1
            for user in self._users.values():
                if not user.isOffLine():
                    books_count = len(list(user.getBooks()))
                    total_steps += books_count
                    if self._provider.supportGroup():
                        total_steps += (books_count * 2)
            self.total = total_steps
        return self.total > total

    def callback(self, success=True, error=None):
        if not success:
            self.status = 'ERROR'
            self.error_cause = error
        elif self.status != 'NO_USER':
            self.status = 'SUCCESS'

        if self._job.cancel.isDisposed():
            if self._listener is not None:
                self._listener.jobFinished(self._job, None)
            return

        if self.status in ('SUCCESS', 'NO_USER') and not self._job.cancel.isSet():
            delay = self._timeout
        else:
            delay = 60

        result = getNamedValue(DeactivationTime=delay)
        if self._listener is not None:
            self._listener.jobFinished(self._job, (result,))

    def _hasConnectedUser(self):
        return any(user.hasSession() for user in self._users.values())
