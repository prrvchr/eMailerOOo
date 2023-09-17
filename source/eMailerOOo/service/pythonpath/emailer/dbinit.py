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

from com.sun.star.sdbc import SQLException

from .unotool import checkVersion
from .unotool import getResourceLocation
from .unotool import getSimpleFile

from .dbconfig import g_csv
from .dbconfig import g_folder
from .dbconfig import g_version

from .dbqueries import getSqlQuery

from .dbtool import createDataSource
from .dbtool import createStaticTable
from .dbtool import executeQueries
from .dbtool import executeSqlQueries
from .dbtool import getDataFromResult
from .dbtool import getDataSourceCall
from .dbtool import getDataSourceConnection
from .dbtool import getSequenceFromResult
from .dbtool import registerDataSource

from .dbtool import checkDataBase
from .dbtool import createStaticTable

import traceback


def getDataSourceUrl(ctx, dbname, plugin, register):
    error = None
    url = getResourceLocation(ctx, plugin, g_folder)
    odb = '%s/%s.odb' % (url, dbname)
    if not getSimpleFile(ctx).exists(odb):
        dbcontext = ctx.ServiceManager.createInstance('com.sun.star.sdb.DatabaseContext')
        datasource = createDataSource(dbcontext, url, dbname)
        error = _createDataBase(ctx, datasource, url, dbname)
        if error is None:
            datasource.DatabaseDocument.storeAsURL(odb, ())
            if register:
                registerDataSource(dbcontext, dbname, odb)
    return url, error

def _createDataBase(ctx, datasource, url, dbname):
    error = None
    print("dbinit._createDataBase() 1")
    try:
        connection = datasource.getConnection('', '')
    except SQLException as e:
        error = e
    else:
        ver, error = checkDataBase(ctx, connection)
        if error is None:
            statement = connection.createStatement()
            createStaticTable(ctx, statement, _getStaticTables(), g_csv)
            tables, queries = _getTablesAndStatements(ctx, statement, ver)
            executeSqlQueries(statement, tables)
            _executeQueries(ctx, statement, _getQueries())
            executeSqlQueries(statement, queries)
            print("dbinit._createDataBase() 2")
            views, triggers = _getViewsAndTriggers(ctx, statement)
            executeSqlQueries(statement, views)
            #executeSqlQueries(statement, triggers)
        connection.close()
        connection.dispose()
    print("dbinit._createDataBase() 3")
    return error

def _executeQueries(ctx, statement, queries):
    for name, format in queries.items():
        query = getSqlQuery(ctx, name, format)
        statement.executeQuery(query)

def _getTableColumns(connection, tables):
    columns = {}
    metadata = connection.MetaData
    for table in tables:
        columns[table] = _getColumns(metadata, table)
    return columns

def _getColumns(metadata, table):
    columns = []
    result = metadata.getColumns("", "", table, "%")
    while result.next():
        column = '"%s"' % result.getString(4)
        columns.append(column)
    return columns

def _createPreparedStatement(ctx, datasource, statements):
    queries = datasource.getQueryDefinitions()
    for name, sql in statements.items():
        if not queries.hasByName(name):
            query = ctx.ServiceManager.createInstance("com.sun.star.sdb.QueryDefinition")
            query.Command = sql
            queries.insertByName(name, query)

def getTablesAndStatements(ctx, connection, version=g_version):
    tables = []
    statements = []
    statement = connection.createStatement()
    query = getSqlQuery(ctx, 'getTableName')
    result = statement.executeQuery(query)
    sequence = getSequenceFromResult(result)
    result.close()
    statement.close()
    call = getDataSourceCall(ctx, connection, 'getTables')
    for table in sequence:
        view = False
        versioned = False
        columns = []
        primary = []
        unique = []
        constraint = {}
        call.setString(1, table)
        result = call.executeQuery()
        while result.next():
            data = getDataFromResult(result)
            view = data.get('View')
            versioned = data.get('Versioned')
            column = data.get('Column')
            definition = '"%s"' % column
            definition += ' %s' % data.get('Type')
            default = data.get('Default')
            definition += ' DEFAULT %s' % default if default else ''
            options = data.get('Options')
            definition += ' %s' % options if options else ''
            columns.append(definition)
            if data.get('Primary'):
                primary.append('"%s"' % column)
            if data.get('Unique'):
                unique.append({'Table': table, 'Column': column})
            if data.get('ForeignTable') and data.get('ForeignColumn'):
                foreign = data.get('ForeignTable')
                if foreign in constraint:
                    constraint[foreign]['ColumnNames'] += column
                    constraint[foreign]['Columns'] += ',"%s"' % column
                    constraint[foreign]['ForeignColumns'] += ',"%s"' % data.get('ForeignColumn')
                else:
                    constraint[foreign] = {'Table': table,
                                           'ColumnNames': column,
                                           'Columns': '"%s"' % column,
                                           'ForeignTable': foreign,
                                           'ForeignColumns': '"%s"' % data.get('ForeignColumn')}
        result.close()
        if primary:
            columns.append(getSqlQuery(ctx, 'getPrimayKey', primary))
        for format in unique:
            columns.append(getSqlQuery(ctx, 'getUniqueConstraint', format))
        for format in constraint.values():
            columns.append(getSqlQuery(ctx, 'getForeignConstraint', format))
        if checkVersion(version, g_version) and versioned:
            columns.append(getSqlQuery(ctx, 'getPeriodColumns'))
        format = (table, ','.join(columns))
        query = getSqlQuery(ctx, 'createTable', format)
        if checkVersion(version, g_version) and versioned:
            query += getSqlQuery(ctx, 'getSystemVersioning')
        tables.append(query)
    call.close()
    return tables, statements

def getViewsAndTriggers(ctx, statement):
    c1 = []
    s1 = []
    f1 = []
    queries = []
    triggers = []
    triggercore = []
    call = getDataSourceCall(ctx, statement.getConnection(), 'getViews')
    tables = getSequenceFromResult(statement.executeQuery(getSqlQuery(ctx, 'getViewName')))
    for table in tables:
        call.setString(1, table)
        result = call.executeQuery()
        while result.next():
            c2 = []
            s2 = []
            f2 = []
            trigger = {}
            data = getDataFromResult(result)
            view = data['View']
            ptable = data['PrimaryTable']
            pcolumn = data['PrimaryColumn']
            labelid = data['LabelId']
            typeid = data['TypeId']
            c1.append('"%s"' % view)
            c2.append('"%s"' % pcolumn)
            c2.append('"Value"')
            s1.append('"%s"."Value"' % view)
            s2.append('"%s"."%s"' % (table, pcolumn))
            s2.append('"%s"."Value"' % table)
            f = 'LEFT JOIN "%s" ON "%s"."%s"="%s"."%s"' % (view, ptable, pcolumn, view, pcolumn)
            f1.append(f)
            f2.append('"%s"' % table)
            f = 'JOIN "Labels" ON "%s"."Label"="Labels"."Label" AND "Labels"."Label"=%s'
            f2.append(f % (table, labelid))
            if typeid is not None:
                f = 'JOIN "Types" ON "%s"."Type"="Types"."Type" AND "Types"."Type"=%s'
                f2.append(f % (table, typeid))
            format = (view, ','.join(c2), ','.join(s2), ' '.join(f2))
            query = getSqlQuery(ctx, 'createView', format)
            queries.append(query)
            triggercore.append(getSqlQuery(ctx, 'createTriggerUpdateAddressBookCore', data))
    call.close()
    if queries:
        column = 'Resource'
        c1.insert(0, '"%s"' % column)
        s1.insert(0, '"%s"."%s"' % (ptable, column))
        f1.insert(0, '"%s"' % ptable)
        f1.append('ORDER BY "%s"."%s"' % (ptable, pcolumn))
        format = ('AddressBook', ','.join(c1), ','.join(s1), ' '.join(f1))
        query = getSqlQuery(ctx, 'createView', format)
        #print("dbinit._getViewsAndTriggers() %s"  % query)
        queries.append(query)
        trigger = getSqlQuery(ctx, 'createTriggerUpdateAddressBook', ' '.join(triggercore))
        triggers.append(trigger)
    return queries, triggers

def getStaticTables():
    tables = ('Tables',
              'Columns',
              'TableColumn',
              'Settings',
              'ConnectionType',
              'AuthenticationType')
    return tables

def getQueries():
    return (('createSpoolerView', None),
            ('createGetServers', None),
            ('createMergeProvider', None),
            ('createMergeDomain', None),
            ('createMergeServer', None),
            ('createUpdateServer', None),
            ('createMergeUser', None),
            ('createInsertJob', None),
            ('createInsertMergeJob', None),
            ('createDeleteJobs', None),
            ('createGetRecipient', None),
            ('createGetMailer', None),
            ('createGetAttachments', None),
            ('createUpdateMailer', None))
