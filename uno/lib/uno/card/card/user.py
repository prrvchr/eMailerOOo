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

from com.sun.star.logging.LogLevel import INFO

from com.sun.star.ucb.ConnectionMode import ONLINE

from .base import UserBase

from .helper import getSqlException
from .helper import getUserId
from .helper import getUserSchema

from ..unotool import getConnectionMode

from ..configuration import g_extension


class User(UserBase):
    def __init__(self, ctx, source, logger, database, provider, scheme, session):
        self._cls = 'User'
        print("User.__init__() 1 url: %s" % session.Url)
        mtd = '__init__'
        logger.logprb(INFO, self._cls, mtd, 1501, session.Name)
        self._ctx = ctx
        logger.logprb(INFO, self._cls, mtd, 1502, session.Name)
        metadata, args = database.selectUser(session.Server, session.Name)
        if metadata is not None:
            session.IsNew = False
            logger.logprb(INFO, self._cls, mtd, 1503, session.Name)
            request = provider.getRequest(session.Url, session.Name)
            if request is None:
                raise getSqlException(ctx, source, 1002, 1506, self._cls, mtd, session.Name, g_extension)
        else:
            session.IsNew = True
            logger.logprb(INFO, self._cls, mtd, 1504, session.Name)
            if getConnectionMode(ctx, session.Server) != ONLINE:
                raise getSqlException(ctx, source, 1004, 1507, self._cls, mtd, session.Server)
            request = provider.getRequest(session.Url, session.Name)
            if request is None:
                raise getSqlException(ctx, source, 1002, 1506, self._cls, mtd, session.Name, g_extension)
            metadata, args = provider.insertUser(database, request, scheme, session.Server, session.Name, session.Token)
            if metadata is None:
                raise getSqlException(ctx, source, 1005, 1508, self._cls, mtd, session.Name)
            database.createUser(getUserSchema(metadata), getUserId(metadata), session.Name, '')
        super().__init__(ctx, request, metadata, args, session)
