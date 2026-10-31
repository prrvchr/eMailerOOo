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

from com.sun.star.sdbcx import CheckOption

from .dbtool import getConnectionUrl
from .dbtool import getDataSourceConnection
from .dbtool import getDriverInfos

from .dbqueries import getSqlQuery

from .dbconfig import g_catalog
from .dbconfig import g_drvinfos
from .dbconfig import g_folder
from .dbconfig import g_schema

from .configuration import g_basename
from .configuration import g_sep

import traceback


def getDataBaseUrl(ctx):
    return getConnectionUrl(ctx, g_folder + g_sep + g_basename)

def getDataSourceUrl(ctx):
    return getConnectionUrl(ctx, g_folder + g_sep + g_basename + '.odb')

def getDataBaseConnection(ctx, url, user, pwd, new=False, infos=None):
    if new:
        infos = getDriverInfos(ctx, url, g_drvinfos)
    return getDataSourceConnection(ctx, url, user, pwd, new, infos)

def getProcedures():
    for name in ('GetBatchJobs', 'GetBatchState', 'AddJob', 'AddMergeJob',
                 'DeleteJobs', 'GetAttachments', 'UpdateSpooler', 'UpdateJobsState',
                 'UpdateJobState', 'ResubmitJobs', 'GetSendJobs', 'GetJobs'):
        yield name

def getViews(ctx, connection):
    yield from _getViews(ctx, g_catalog, g_schema, 'Spooler', CheckOption.CASCADE)

def _getViews(ctx, catalog, schema, view, *option):
    format = {'Catalog': catalog, 'Schema': schema}
    yield catalog, schema, view, getSqlQuery(ctx, 'get%sViewCommand' % view, format), *option

