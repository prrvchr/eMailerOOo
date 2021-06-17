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

# DataBase creation Queries
    # Create Text Table Queries
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

    # Create Text Table Options
    elif name == 'setTableSource':
        query = 'SET TABLE "%s" SOURCE "%s"' % (format, g_csv % format)
    elif name == 'setTableHeader':
        query = 'SET TABLE "%s" SOURCE HEADER "%s"' % format
    elif name == 'setTableReadOnly':
        query = 'SET TABLE "%s" READONLY TRUE' % format

    # Create Cached Table Queries
    elif name == 'createTable':
        query = 'CREATE CACHED TABLE IF NOT EXISTS "%s"(%s)' % format

    # Create Cached Table Options
    elif name == 'getPrimayKey':
        query = 'PRIMARY KEY(%s)' % ','.join(format)

    elif name == 'getUniqueConstraint':
        query = 'CONSTRAINT "Unique%(Table)s%(Column)s" UNIQUE("%(Column)s")' % format

    elif name == 'getForeignConstraint':
        q = 'CONSTRAINT "Foreign%(Table)s%(ColumnNames)s" FOREIGN KEY(%(Columns)s) REFERENCES '
        q += '"%(ForeignTable)s"(%(ForeignColumns)s) ON DELETE CASCADE ON UPDATE CASCADE'
        query = q % format

    elif name == 'getPeriodColumns':
        query = '"RowStart" TIMESTAMP GENERATED ALWAYS AS ROW START,'
        query += '"RowEnd" TIMESTAMP GENERATED ALWAYS AS ROW END,'
        query += 'PERIOD FOR SYSTEM_TIME("RowStart","RowEnd")'

    elif name == 'getSystemVersioning':
        query = ' WITH SYSTEM VERSIONING'

    # Create View Queries
    elif name == 'createSpoolerView':
        c1 = '"JobId"'
        c2 = '"BatchId"'
        c3 = '"State"'
        c4 = '"Subject"'
        c5 = '"Sender"'
        c6 = '"Recipient"'
        c7 = '"Document"'
        c8 = '"DataSource"'
        c9 = '"Query"'
        c10 = '"Submit"'
        c11 = '"Sending"'
        c = (c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11)
        s1 = '"Recipients"."JobId"'
        s2 = '"Senders"."BatchId"'
        s3 = '"Recipients"."State"'
        s4 = '"Senders"."Subject"'
        s5 = '"Senders"."Sender"'
        s6 = '"Recipients"."Recipient"'
        s7 = '"Senders"."Document"'
        s8 = '"Senders"."DataSource"'
        s9 = '"Senders"."Query"'
        s10 = '"Senders"."TimeStamp"'
        s11 = '"Recipients"."TimeStamp"'
        s = (s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11)
        f1 = '"Senders"'
        f2 = 'JOIN "Recipients" ON "Senders"."BatchId"="Recipients"."BatchId"'
        f = (f1,f2)
        p = (','.join(c), ','.join(s), ' '.join(f))
        query = 'CREATE VIEW "Spooler" (%s) AS SELECT %s FROM %s;' % p

# Select Queries
    # DataBase creation Select Queries
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

    # IspDb Select Queries
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

    # Mail Select Queries
    elif name == 'getSenders':
        query = 'SELECT "User" FROM "Users" ORDER BY "TimeStamp";'

    # Merger Composer Select Queries
    elif name == 'getQueryCommand':
        query = 'SELECT * FROM "%s"' % format

    elif name == 'getRecipientColumns':
        query = '"%s"' % '", "'.join(format)
        if len(format) > 1:
            query = 'COALESCE(%s)' % query

    elif name == 'getRecipientQuery':
        query = 'SELECT %s AS "Recipient", "%s" AS "Identifier" FROM "%s" WHERE %s;' % format

    # Spooler Select Queries
    elif name == 'getViewQuery':
        query = 'SELECT * FROM "Spooler";'

    elif name == 'getSpoolerViewQuery':
        query = 'SELECT * FROM "View";'

    # MailSpooler Select Queries
    elif name == 'getSpoolerJobs':
        query = 'SELECT "JobId" FROM "Spooler" WHERE "State" = ? ORDER BY "JobId";'

# Delete Queries
    # Mail Delete Queries
    elif name == 'deleteUser':
        query = 'DELETE FROM "Users" WHERE "User" = ?;'

# Update Queries
    # MailSpooler Update Queries
    elif name == 'setJobState':
        query = 'UPDATE "Recipients" SET "State"=? WHERE "JobId"=?;'

# Function creation Queries
    # IspDb Function Queries
    elif name == 'createGetDomain':
        query = """\
CREATE FUNCTION "GetDomain"("User" VARCHAR(320))
  RETURNS VARCHAR(100)
  SPECIFIC "GetDomain_1"
  RETURN SUBSTRING("User" FROM POSITION('@' IN "User") + 1);
"""

# Select Procedure Queries
    # IspDb Select Procedure Queries
    elif name == 'createGetServers':
        query = """\
CREATE PROCEDURE "GetServers"(IN "Email" VARCHAR(320),
                              IN "DomainName" VARCHAR(255),
                              OUT "Server" VARCHAR(255),
                              OUT "Port" SMALLINT,
                              OUT "LoginName" VARCHAR(100),
                              OUT "Password" VARCHAR(100))
  SPECIFIC "GetServers_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE "Result" CURSOR WITH RETURN FOR
      SELECT "Servers"."Server", "Servers"."Port", "Servers"."Connection",
      "Servers"."Authentication", "Servers"."LoginMode" FROM "Servers"
      JOIN "Providers" ON "Servers"."Provider"="Providers"."Provider"
      LEFT JOIN "Domains" ON "Providers"."Provider"="Domains"."Provider"
      WHERE "Providers"."Provider"="DomainName" OR "Domains"."Domain"="DomainName"
      FOR READ ONLY;
    SET ("Server", "Port", "LoginName", "Password") = (SELECT "Users"."Server",
     "Users"."Port", "Users"."LoginName", "Users"."Password" FROM "Users"
      WHERE "Users"."User"="Email");
    OPEN "Result";
  END;"""

    # MailSpooler Select Procedure Queries
    elif name == 'createGetRecipient':
        query = """\
CREATE PROCEDURE "GetRecipient"(IN "Id" INTEGER)
  SPECIFIC "GetRecipient_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE "Result" CURSOR WITH RETURN FOR
      SELECT "Recipient", "Identifier", "BatchId" From "Recipients"
      WHERE "JobId"="Id"
      FOR READ ONLY;
    OPEN "Result";
  END;"""

    elif name == 'createGetSender':
        query = """\
CREATE PROCEDURE "GetSender"(IN "Id" INTEGER)
  SPECIFIC "GetSender_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE "Result" CURSOR WITH RETURN FOR
      SELECT "Sender", "Subject", "Document", "DataSource", "Query" FROM "Senders"
      WHERE "BatchId"="Id"
      FOR READ ONLY;
    OPEN "Result";
  END;"""

    elif name == 'createGetAttachments':
        query = """\
CREATE PROCEDURE "GetAttachments"(IN "Id" INTEGER)
  SPECIFIC "GetAttachments_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE "Result" CURSOR WITH RETURN FOR
      SELECT "Attachment" From "Attachments"
      WHERE "BatchId"="Id" ORDER BY "TimeStamp"
      FOR READ ONLY;
    OPEN "Result";
  END;"""

    elif name == 'createGetServer':
        query = """\
CREATE PROCEDURE "GetServer"(IN "User" VARCHAR(320))
  SPECIFIC "GetServer_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE "Result" CURSOR WITH RETURN FOR
      SELECT "Servers"."Server", "Servers"."Port", "Servers"."Connection",
      "Servers"."Authentication", "Servers"."LoginMode",
      "Users"."LoginName", "Users"."Password"
      FROM "Servers"
      JOIN "Users" ON "Servers"."Server"="Users"."Server" AND "Servers"."Port"="Users"."Port"
      WHERE "Users"."User"="User"
      FOR READ ONLY;
    OPEN "Result";
  END;"""

    elif name == 'createGetJobMail1':
        query = """\
CREATE PROCEDURE "GetJobMail"(IN "Job" INTEGER)
  SPECIFIC "GetJobMail_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE "Result" CURSOR WITH RETURN FOR
      SELECT "Senders"."Sender", "Senders"."Subject", "Senders"."Document",
      "Senders"."DataSource", "Senders"."Query", "Recipients"."Recipient",
      "Recipients"."Identifier", ARRAY_AGG("Attachments"."Attachment")
      FROM "Senders"
      JOIN "Recipients" ON "Senders"."BatchId"="Recipients"."BatchId"
      LEFTJOIN "Attachments" ON ""Senders"."BatchId"="Attachments"."BatchId"
      WHERE "Recipients"."JobId"="Job"
      GROUP BY "Senders"."Sender", "Senders"."Subject", "Senders"."Document",
      "Senders"."DataSource", "Senders"."Query", "Recipients"."Recipient",
      "Recipients"."Identifier"
      FOR READ ONLY;
    OPEN "Result";
  END;"""

# Insert Procedure Queries
    # MailServiceSpooler Insert Procedure Queries
    elif name == 'createInsertJob':
        query = """\
CREATE PROCEDURE "InsertJob"(IN "Sender" VARCHAR(320),
                             IN "Subject" VARCHAR(78),
                             IN "Document" VARCHAR(512),
                             IN "Recipients" VARCHAR(320) ARRAY,
                             IN "Attachments" VARCHAR(512) ARRAY,
                             OUT "Id" INTEGER)
  SPECIFIC "InsertJob_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    DECLARE "BatchId" INTEGER DEFAULT 0;
    DECLARE "Index" INTEGER DEFAULT 1;
    INSERT INTO "Senders" ("Sender","Subject","Document")
    VALUES ("Sender","Subject","Document");
    SET "BatchId" = IDENTITY();
    WHILE "Index" <= CARDINALITY("Recipients") DO
      INSERT INTO "Recipients" ("BatchId","Recipient")
      VALUES ("BatchId","Recipients"["Index"]);
      SET "Index" = "Index" + 1;
    END WHILE;
    SET "Index" = 1;
    WHILE "Index" <= CARDINALITY("Attachments") DO
      INSERT INTO "Attachments" ("BatchId","Attachment")
      VALUES ("BatchId","Attachments"["Index"]);
      SET "Index" = "Index" + 1;
    END WHILE;
    SET "Id" = "BatchId";
  END;"""

    # MailServiceSpooler Insert Procedure Queries
    elif name == 'createInsertMergeJob':
        query = """\
CREATE PROCEDURE "InsertMergeJob"(IN "Sender" VARCHAR(320),
                                  IN "Subject" VARCHAR(78),
                                  IN "Document" VARCHAR(512),
                                  IN "DataSource" VARCHAR(512),
                                  IN "Query" VARCHAR(512),
                                  IN "Recipients" VARCHAR(320) ARRAY,
                                  IN "Identifiers" VARCHAR(128) ARRAY,
                                  IN "Attachments" VARCHAR(512) ARRAY,
                                  OUT "Id" INTEGER)
  SPECIFIC "InsertMergeJob_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    DECLARE "BatchId" INTEGER DEFAULT 0;
    DECLARE "Index" INTEGER DEFAULT 1;
    INSERT INTO "Senders" ("Sender","Subject","Document","DataSource","Query")
    VALUES ("Sender","Subject","Document", "DataSource","Query");
    SET "BatchId" = IDENTITY();
    WHILE "Index" <= CARDINALITY("Recipients") DO
      INSERT INTO "Recipients" ("BatchId","Recipient","Identifier")
      VALUES ("BatchId","Recipients"["Index"],"Identifiers"["Index"]);
      SET "Index" = "Index" + 1;
    END WHILE;
    SET "Index" = 1;
    WHILE "Index" <= CARDINALITY("Attachments") DO
      INSERT INTO "Attachments" ("BatchId","Attachment")
      VALUES ("BatchId","Attachments"["Index"]);
      SET "Index" = "Index" + 1;
    END WHILE;
    SET "Id" = "BatchId";
  END;"""

# Update Procedure Queries
    # IspDb Update Procedure Queries
    elif name == 'createUpdateServer':
        query = """\
CREATE PROCEDURE "UpdateServer"(IN "Server1" VARCHAR(255),
                                IN "Port1" SMALLINT,
                                IN "Server2" VARCHAR(255),
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

# Merge Procedure Queries
    # IspDb Merge Procedure Queries
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

    # IspDb Merge Procedure Queries
    elif name == 'createMergeDomain':
        query = """\
CREATE PROCEDURE "MergeDomain"(IN "Provider" VARCHAR(100),
                               IN "Domain" VARCHAR(255),
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

    # IspDb Merge Procedure Queries
    elif name == 'createMergeServer':
        query = """\
CREATE PROCEDURE "MergeServer"(IN "Provider" VARCHAR(100),
                               IN "Server" VARCHAR(255),
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

    # IspDb Merge Procedure Queries
    elif name == 'createMergeUser':
        query = """\
CREATE PROCEDURE "MergeUser"(IN "User" VARCHAR(320),
                             IN "Server" VARCHAR(255),
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

# Call Procedure Query
    elif name == 'getServers':
        query = 'CALL "GetServers"(?,?,?,?,?,?)'
    elif name == 'getRecipient':
        query = 'CALL "GetRecipient"(?)'
    elif name == 'getSender':
        query = 'CALL "GetSender"(?)'
    elif name == 'getAttachments':
        query = 'CALL "GetAttachments"(?)'
    elif name == 'getServer':
        query = 'CALL "GetServer"(?)'
    elif name == 'insertJob':
        query = 'CALL "InsertJob"(?,?,?,?,?,?)'
    elif name == 'insertMergeJob':
        query = 'CALL "InsertMergeJob"(?,?,?,?,?,?,?,?,?)'
    elif name == 'updateServer':
        query = 'CALL "UpdateServer"(?,?,?,?,?,?,?)'
    elif name == 'mergeProvider':
        query = 'CALL "MergeProvider"(?,?,?,?)'
    elif name == 'mergeDomain':
        query = 'CALL "MergeDomain"(?,?,?)'
    elif name == 'mergeServer':
        query = 'CALL "MergeServer"(?,?,?,?,?,?,?)'
    elif name == 'mergeUser':
        query = 'CALL "MergeUser"(?,?,?,?,?,?)'

# ShutDown Queries
    # Normal ShutDown Queries
    elif name == 'shutdown':
        query = 'SHUTDOWN;'
    # Compact ShutDown Queries
    elif name == 'shutdownCompact':
        query = 'SHUTDOWN COMPACT;'

# Get prepareCommand Query
    elif name == 'prepareCommand':
        query = 'SELECT * FROM "%s"' % format

# Get Users and Privileges Query
    elif name == 'getUsers':
        query = 'SELECT * FROM INFORMATION_SCHEMA.SYSTEM_USERS'
    elif name == 'getPrivileges':
        query = 'SELECT * FROM INFORMATION_SCHEMA.TABLE_PRIVILEGES'
    elif name == 'changePassword':
        query = "SET PASSWORD '%s'" % format

# Queries don't exist!!!
    else:
        query = None
        msg = getMessage(ctx, g_message, 101, name)
        logMessage(ctx, SEVERE, msg, 'dbqueries', 'getSqlQuery()')
    return query
