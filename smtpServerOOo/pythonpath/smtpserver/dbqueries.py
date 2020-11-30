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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .dbconfig import g_csv

from .logger import logMessage
from .logger import getMessage
g_message = 'dbqueries'


def getSqlQuery(ctx, name, format=None):

# Create Static Table Queries
    if name == 'createTableTables':
        c1 = '"Table" INTEGER NOT NULL PRIMARY KEY'
        c2 = '"Name" VARCHAR(100) NOT NULL'
        c3 = '"Identity" INTEGER DEFAULT NULL'
        c4 = '"View" BOOLEAN DEFAULT TRUE'
        c5 = '"Versioned" BOOLEAN DEFAULT FALSE'
        k1 = 'CONSTRAINT "UniqueTablesName" UNIQUE("Name")'
        c = (c1, c2, c3, c4, c5, k1)
        query = 'CREATE TEXT TABLE IF NOT EXISTS "Tables"(%s);' % ','.join(c)
    elif name == 'createTableColumns':
        c1 = '"Column" INTEGER NOT NULL PRIMARY KEY'
        c2 = '"Name" VARCHAR(100) NOT NULL'
        k1 = 'CONSTRAINT "UniqueColumnsName" UNIQUE("Name")'
        c = (c1, c2, k1)
        query = 'CREATE TEXT TABLE IF NOT EXISTS "Columns"(%s);' % ','.join(c)
    elif name == 'createTableTableColumn':
        c1 = '"Table" INTEGER NOT NULL'
        c2 = '"Column" INTEGER NOT NULL'
        c3 = '"TypeName" VARCHAR(100) NOT NULL'
        c4 = '"TypeLenght" SMALLINT DEFAULT NULL'
        c5 = '"Default" VARCHAR(100) DEFAULT NULL'
        c6 = '"Options" VARCHAR(100) DEFAULT NULL'
        c7 = '"Primary" BOOLEAN NOT NULL'
        c8 = '"Unique" BOOLEAN NOT NULL'
        c9 = '"ForeignTable" INTEGER DEFAULT NULL'
        c10 = '"ForeignColumn" INTEGER DEFAULT NULL'
        k1 = 'PRIMARY KEY("Table","Column")'
        k2 = 'CONSTRAINT "ForeignTableColumnTable" FOREIGN KEY("Table") REFERENCES '
        k2 += '"Tables"("Table") ON DELETE CASCADE ON UPDATE CASCADE'
        k3 = 'CONSTRAINT "ForeignTableColumnColumn" FOREIGN KEY("Column") REFERENCES '
        k3 += '"Columns"("Column") ON DELETE CASCADE ON UPDATE CASCADE'
        c = (c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, k1, k2, k3)
        query = 'CREATE TEXT TABLE IF NOT EXISTS "TableColumn"(%s);' % ','.join(c)
    elif name == 'createTableSettings':
        c1 = '"Id" INTEGER NOT NULL PRIMARY KEY'
        c2 = '"Name" VARCHAR(100) NOT NULL'
        c3 = '"Value1" VARCHAR(100) NOT NULL'
        c4 = '"Value2" VARCHAR(100) DEFAULT NULL'
        c5 = '"Value3" VARCHAR(100) DEFAULT NULL'
        c = (c1, c2, c3, c4, c5)
        p = ','.join(c)
        query = 'CREATE TEXT TABLE IF NOT EXISTS "Settings"(%s);' % p
    elif name == 'setTableSource':
        query = 'SET TABLE "%s" SOURCE "%s"' % (format, g_csv % format)
    elif name == 'setTableHeader':
        query = 'SET TABLE "%s" SOURCE HEADER "%s"' % format
    elif name == 'setTableReadOnly':
        query = 'SET TABLE "%s" READONLY TRUE' % format

# Create Cached Table Options
    elif name == 'getPrimayKey':
        query = 'PRIMARY KEY(%s)' % ','.join(format)

    elif name == 'getUniqueConstraint':
        query = 'CONSTRAINT "Unique%(Table)s%(Column)s" UNIQUE("%(Column)s")' % format

    elif name == 'getForeignConstraint':
        q = 'CONSTRAINT "Foreign%(Table)s%(ColumnNames)s" FOREIGN KEY(%(Columns)s) REFERENCES '
        q += '"%(ForeignTable)s"(%(ForeignColumns)s) ON DELETE CASCADE ON UPDATE CASCADE'
        query = q % format

# Create Cached Table Queries
    elif name == 'createTable':
        query = 'CREATE CACHED TABLE IF NOT EXISTS "%s"(%s)' % format

    elif name == 'getPeriodColumns':
        query = '"RowStart" TIMESTAMP GENERATED ALWAYS AS ROW START,'
        query += '"RowEnd" TIMESTAMP GENERATED ALWAYS AS ROW END,'
        query += 'PERIOD FOR SYSTEM_TIME("RowStart","RowEnd")'

    elif name == 'getSystemVersioning':
        query = ' WITH SYSTEM VERSIONING'

# Select Queries
    elif name == 'getTableName':
        query = 'SELECT "Name" FROM "Tables" ORDER BY "Table";'
    elif name == 'getTables':
        s1 = '"T"."Table" AS "TableId"'
        s2 = '"C"."Column" AS "ColumnId"'
        s3 = '"T"."Name" AS "Table"'
        s4 = '"C"."Name" AS "Column"'
        s5 = '"TC"."TypeName" AS "Type"'
        s6 = '"TC"."TypeLenght" AS "Lenght"'
        s7 = '"TC"."Default"'
        s8 = '"TC"."Options"'
        s9 = '"TC"."Primary"'
        s10 = '"TC"."Unique"'
        s11 = '"TC"."ForeignTable" AS "ForeignTableId"'
        s12 = '"TC"."ForeignColumn" AS "ForeignColumnId"'
        s13 = '"T2"."Name" AS "ForeignTable"'
        s14 = '"C2"."Name" AS "ForeignColumn"'
        s15 = '"T"."View"'
        s16 = '"T"."Versioned"'
        s = (s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12,s13,s14,s15,s16)
        f1 = '"Tables" AS "T"'
        f2 = 'JOIN "TableColumn" AS "TC" ON "T"."Table"="TC"."Table"'
        f3 = 'JOIN "Columns" AS "C" ON "TC"."Column"="C"."Column"'
        f4 = 'LEFT JOIN "Tables" AS "T2" ON "TC"."ForeignTable"="T2"."Table"'
        f5 = 'LEFT JOIN "Columns" AS "C2" ON "TC"."ForeignColumn"="C2"."Column"'
        w = '"T"."Name"=?'
        f = (f1, f2, f3, f4, f5)
        p = (','.join(s), ' '.join(f), w)
        query = 'SELECT %s FROM %s WHERE %s;' % p
    elif name == 'getTablesName':
        query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.SYSTEM_TABLES WHERE TABLE_TYPE='TABLE'"

    elif name == 'getSmtpServers':
        s1 = '"Servers"."Server"'
        s2 = '"Servers"."Port"'
        s3 = '"Servers"."Connection"'
        s4 = '"Servers"."Authentication"'
        s5 = '"Servers"."LoginMode"'
        s = (s1,s2,s3,s4,s5)
        f1 = '"Servers"'
        f2 = 'JOIN "Providers" ON "Servers"."Provider"="Providers"."Provider"'
        f3 = 'LEFT JOIN "Domains" ON "Providers"."Provider"="Domains"."Provider"'
        f = (f1, f2, f3)
        w = '"Providers"."Provider"=? OR "Domains"."Domain"=?'
        p = (','.join(s), ' '.join(f), w)
        query = 'SELECT %s FROM %s WHERE %s;' % p

# Function creation Queries
    elif name == 'createGetDomain':
        query = """\
CREATE FUNCTION "GetDomain"("User" VARCHAR(100))
  RETURNS VARCHAR(100)
  SPECIFIC "GetDomain_1"
  RETURN SUBSTRING("User" FROM POSITION('@' IN "User") + 1);
"""

# Select Procedure Queries
    elif name == 'createGetUser':
        query = """\
CREATE PROCEDURE "GetUser"(IN "User" VARCHAR(100),
                           OUT "Server" VARCHAR(100),
                           OUT "Port" SMALLINT,
                           OUT "LoginName" VARCHAR(100),
                           OUT "Password" VARCHAR(100),
                           OUT "Domain" VARCHAR(100))
  SPECIFIC "GetUser_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE "Result" CURSOR WITH RETURN FOR
      SELECT "Servers"."Server", "Servers"."Port", "Servers"."Connection",
      "Servers"."Authentication", "Servers"."LoginMode" FROM "Servers"
      JOIN "Providers" ON "Servers"."Provider"="Providers"."Provider"
      LEFT JOIN "Domains" ON "Providers"."Provider"="Domains"."Provider"
      WHERE "Providers"."Provider"="GetDomain"("User") OR "Domains"."Domain"="GetDomain"("User")
      FOR READ ONLY;
    SET ("Server", "Port", "LoginName", "Password") = (SELECT "Server", "Port",
    "LoginName", "Password" FROM "Users" WHERE "Users"."User"="User");
    SET "Domain" = "GetDomain"("User");
    OPEN "Result";
  END;"""

# Merge Procedure Queries
    elif name == 'createMergeProvider':
        query = """\
CREATE PROCEDURE "MergeProvider"(IN "Provider" VARCHAR(100),
                                 IN "DisplayName"  VARCHAR(100),
                                 IN "DisplayShortName" VARCHAR(100),
                                 IN "Time" TIMESTAMP(6))
  SPECIFIC "MergeProvider_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    MERGE INTO "Providers" USING (VALUES("Provider","DisplayName","DisplayShortName","Time"))
      AS vals(w,x,y,z) ON "Providers"."Provider"=vals.w
        WHEN MATCHED THEN UPDATE
          SET "DisplayName"=vals.x, "DisplayShortName"=vals.y, "TimeStamp"=vals.z
        WHEN NOT MATCHED THEN INSERT ("Provider","DisplayName","DisplayShortName","TimeStamp")
          VALUES vals.w, vals.x, vals.y, vals.z;
  END"""
    elif name == 'createMergeDomain':
        query = """\
CREATE PROCEDURE "MergeDomain"(IN "Provider" VARCHAR(100),
                               IN "Domain" VARCHAR(100),
                               IN "Time" TIMESTAMP(6))
  SPECIFIC "MergeDomain_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    MERGE INTO "Domains" USING (VALUES("Domain","Provider","Time"))
      AS vals(x,y,z) ON "Domains"."Domain"=vals.x
        WHEN MATCHED THEN UPDATE
          SET "Provider"=vals.y, "TimeStamp"=vals.z
        WHEN NOT MATCHED THEN INSERT ("Domain","Provider","TimeStamp")
          VALUES vals.x, vals.y, vals.z;
  END"""
    elif name == 'createMergeServer':
        query = """\
CREATE PROCEDURE "MergeServer"(IN "Provider" VARCHAR(100),
                               IN "Server" VARCHAR(100),
                               IN "Port" SMALLINT,
                               IN "Connection" TINYINT,
                               IN "Authentication" TINYINT,
                               IN "LoginMode" TINYINT,
                               IN "Time" TIMESTAMP(6))
  SPECIFIC "MergeServer_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    MERGE INTO "Servers" USING (VALUES("Server","Port","Provider","Connection","Authentication","LoginMode","Time"))
      AS vals(t,u,v,w,x,y,z) ON "Servers"."Server"=vals.t AND "Servers"."Port"=vals.u
        WHEN MATCHED THEN UPDATE
          SET "Provider"=vals.v, "Connection"=vals.w, "Authentication"=vals.x, "LoginMode"=vals.y, "TimeStamp"=vals.z
        WHEN NOT MATCHED THEN INSERT ("Server","Port","Provider","Connection","Authentication","LoginMode","TimeStamp")
          VALUES vals.t, vals.u, vals.v, vals.w, vals.x, vals.y, vals.z;
  END"""
    elif name == 'createUpdateServer':
        query = """\
CREATE PROCEDURE "UpdateServer"(IN "Server1" VARCHAR(100),
                                IN "Port1" SMALLINT,
                                IN "Server2" VARCHAR(100),
                                IN "Port2" SMALLINT,
                                IN "Connection" TINYINT,
                                IN "Authentication" TINYINT,
                                IN "Time" TIMESTAMP(6))
  SPECIFIC "UpdateServer_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    UPDATE "Servers" SET "Server"="Server2", "Port"="Port2",
      "Connection"="Connection", "Authentication"="Authentication",
      "TimeStamp"="Time"
      WHERE "Server" = "Server1" AND "Port" = "Port1";
  END"""
    elif name == 'createMergeUser':
        query = """\
CREATE PROCEDURE "MergeUser"(IN "User" VARCHAR(100),
                             IN "Server" VARCHAR(100),
                             IN "Port" SMALLINT,
                             IN "LoginName" VARCHAR(100),
                             IN "Password" VARCHAR(100),
                             IN "Time" TIMESTAMP(6))
  SPECIFIC "MergeUser_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    MERGE INTO "Users" USING (VALUES("User","Server","Port","LoginName","Password","Time"))
      AS vals(u,v,w,x,y,z) ON "Users"."User"=vals.u
        WHEN MATCHED THEN UPDATE
          SET "Server"=vals.v, "Port"=vals.w, "LoginName"=vals.x, "Password"=vals.y, "TimeStamp"=vals.z
        WHEN NOT MATCHED THEN INSERT ("User","Server","Port","LoginName","Password","TimeStamp")
          VALUES vals.u, vals.v, vals.w, vals.x, vals.y, vals.z;
  END"""

# Get DataBase Version Query
    elif name == 'getVersion':
        query = 'SELECT DISTINCT DATABASE_VERSION() AS "HSQL Version" FROM INFORMATION_SCHEMA.SYSTEM_TABLES'

# ShutDown Queries
    elif name == 'shutdown':
        query = 'SHUTDOWN;'
    elif name == 'shutdownCompact':
        query = 'SHUTDOWN COMPACT;'

# Get Procedure Query
    elif name == 'getUser':
        query = 'CALL "GetUser"(?,?,?,?,?,?)'
    elif name == 'mergeProvider':
        query = 'CALL "MergeProvider"(?,?,?,?)'
    elif name == 'mergeDomain':
        query = 'CALL "MergeDomain"(?,?,?)'
    elif name == 'mergeServer':
        query = 'CALL "MergeServer"(?,?,?,?,?,?,?)'
    elif name == 'updateServer':
        query = 'CALL "UpdateServer"(?,?,?,?,?,?,?)'
    elif name == 'mergeUser':
        query = 'CALL "MergeUser"(?,?,?,?,?,?)'

# Queries don't exist!!!
    else:
        query = None
        msg = getMessage(ctx, g_message, 101, name)
        logMessage(ctx, SEVERE, msg, 'dbqueries', 'getSqlQuery()')
    return query
