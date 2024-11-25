#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-24 https://prrvchr.github.io                                  ║
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

from com.sun.star.sdbc import SQLException

from com.sun.star.sdbcx import CheckOption

from .dbtool import createStaticTables
from .dbtool import createStaticIndexes
from .dbtool import createStaticForeignKeys
from .dbtool import createTables
from .dbtool import createIndexes
from .dbtool import createForeignKeys
from .dbtool import createViews
from .dbtool import executeQueries
from .dbtool import getConnectionInfos
from .dbtool import getDataBaseTables
from .dbtool import getDataBaseIndexes
from .dbtool import getDataBaseForeignKeys
from .dbtool import getDataSourceConnection
from .dbtool import getDriverInfos
from .dbtool import getTableNames
from .dbtool import getTables
from .dbtool import getIndexes
from .dbtool import getForeignKeys
from .dbtool import setStaticTable

from .dbqueries import getSqlQuery

from .dbconfig import g_catalog
from .dbconfig import g_schema
from .dbconfig import g_csv
from .dbconfig import g_drvinfos

from time import sleep
import traceback


def getDataBaseConnection(ctx, url, user, pwd, new=False, infos=None):
    if new:
        infos = getDriverInfos(ctx, url, g_drvinfos)
    return getDataSourceConnection(ctx, url, user, pwd, new, infos)

def createDataBase(ctx, connection, odb):
    print("smtpMailer.DataBase._initialize() 1")
    sleep(0.2)
    tables = connection.getTables()
    statement = connection.createStatement()
    statics = createStaticTables(g_catalog, g_schema, tables)
    createStaticIndexes(g_catalog, g_schema, tables)
    createStaticForeignKeys(g_catalog, g_schema, tables)
    setStaticTable(statement, statics, g_csv, True)
    _createTables(connection, statement, tables)
    _createIndexes(statement, tables)
    _createForeignKeys(statement, tables)
    views = _getViews(ctx, g_catalog, g_schema, 'Spooler', CheckOption.CASCADE)
    createViews(connection.getViews(), views)
    executeQueries(ctx, statement, _getProcedures(), 'create%s')
    statement.close()
    connection.getParent().DatabaseDocument.storeAsURL(odb, ())
    connection.close()
    print("smtpMailer.DataBase._initialize() 2")

def _createTables(connection, statement, tables):
    infos = getConnectionInfos(connection, 'AutoIncrementCreation', 'RowVersionCreation')
    createTables(tables, getDataBaseTables(connection, statement, getTables(), getTableNames(), infos[0], infos[1]))

def _createIndexes(statement, tables):
    createIndexes(tables, getDataBaseIndexes(statement, getIndexes()))

def _createForeignKeys(statement, tables):
    createForeignKeys(tables, getDataBaseForeignKeys(statement, getForeignKeys()))

def _getViews(ctx, catalog, schema, view, *option):
    format = {'Catalog': catalog, 'Schema': schema}
    yield catalog, schema, view, getSqlQuery(ctx, 'get%sViewCommand' % view, format), *option

def _getProcedures():
    for name in ('InsertJob', 'InsertMergeJob', 'DeleteJobs', 'GetRecipient',
                 'GetMailer', 'GetAttachments', 'UpdateSpooler'):
        yield name

