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
        c3 = '"Type" VARCHAR(100) NOT NULL'
        c4 = '"Default" VARCHAR(100) DEFAULT NULL'
        c5 = '"Options" VARCHAR(100) DEFAULT NULL'
        c6 = '"Primary" BOOLEAN NOT NULL'
        c7 = '"Unique" BOOLEAN NOT NULL'
        c8 = '"ForeignTable" INTEGER DEFAULT NULL'
        c9 = '"ForeignColumn" INTEGER DEFAULT NULL'
        k1 = 'PRIMARY KEY("Table","Column")'
        k2 = 'CONSTRAINT "ForeignTableColumnTable" FOREIGN KEY("Table") REFERENCES '
        k2 += '"Tables"("Table") ON DELETE CASCADE ON UPDATE CASCADE'
        k3 = 'CONSTRAINT "ForeignTableColumnColumn" FOREIGN KEY("Column") REFERENCES '
        k3 += '"Columns"("Column") ON DELETE CASCADE ON UPDATE CASCADE'
        c = (c1, c2, c3, c4, c5, c6, c7, c8, c9, k1, k2, k3)
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

    elif name == 'createTableConnectionType':
        c1 = '"Type" INTEGER NOT NULL PRIMARY KEY'
        c2 = '"Connection" VARCHAR(20) NOT NULL'
        c = (c1, c2)
        p = ','.join(c)
        query = 'CREATE TEXT TABLE IF NOT EXISTS "ConnectionType"(%s);' % p

    elif name == 'createTableAuthenticationType':
        c1 = '"Type" INTEGER NOT NULL PRIMARY KEY'
        c2 = '"Authentication" VARCHAR(20) NOT NULL'
        c = (c1, c2)
        p = ','.join(c)
        query = 'CREATE TEXT TABLE IF NOT EXISTS "AuthenticationType"(%s);' % p

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
        s10 = '"Senders"."Created"'
        s11 = '"Recipients"."Created"'
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
        s5 = '"TC"."Type"'
        s6 = '"TC"."Default"'
        s7 = '"TC"."Options"'
        s8 = '"TC"."Primary"'
        s9 = '"TC"."Unique"'
        s10 = '"TC"."ForeignTable" AS "ForeignTableId"'
        s11 = '"TC"."ForeignColumn" AS "ForeignColumnId"'
        s12 = '"T2"."Name" AS "ForeignTable"'
        s13 = '"C2"."Name" AS "ForeignColumn"'
        s14 = '"T"."View"'
        s15 = '"T"."Versioned"'
        s = (s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12,s13,s14,s15)
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
        query = 'SELECT "User" FROM "Users" ORDER BY "Created";'

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

    elif name == 'getBookmark':
        query = 'SELECT "%s" FROM "%s" WHERE "%s" = ?;' % format

    elif name == 'getJobState':
        query = 'SELECT "State" FROM "Recipients" WHERE "JobId" = ?;'

    elif name == 'getJobIds':
        query = 'SELECT ARRAY_AGG("JobId") FROM "Recipients" WHERE "BatchId" = ?;'

# Delete Queries
    # Mail Delete Queries
    elif name == 'deleteUser':
        query = 'DELETE FROM "Users" WHERE "User" = ?;'

# Update Queries
    # MailSpooler Update Queries
    elif name == 'updateRecipient':
        query = 'UPDATE "Recipients" SET "State"=?, "MessageId"=?, "Modified"=DEFAULT WHERE "JobId"=?;'

    elif name == 'setBatchState':
        query = 'UPDATE "Recipients" SET "State"=?, "Modified"=? WHERE "BatchId"=?;'

# Function creation Queries
    # IspDb Function Queries
    elif name == 'createGetDomain':
        query = """\
CREATE FUNCTION "GetDomain"(EMAIL VARCHAR(320))
  RETURNS VARCHAR(100)
  SPECIFIC "GetDomain_1"
  RETURN SUBSTRING(EMAIL FROM POSITION('@' IN EMAIL) + 1);
"""

# Delete Procedure Queries
    # SpoolerService Delete Procedure Queries
    elif name == 'createDeleteJobs':
        query = """\
CREATE PROCEDURE "DeleteJobs"(IN JOBIDS INTEGER ARRAY)
  SPECIFIC "DeleteJobs_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    DELETE FROM "Recipients" WHERE "JobId" IN (UNNEST (JOBIDS));
    DELETE FROM "Senders" WHERE "BatchId" IN (SELECT "BatchId" FROM "Senders"
    EXCEPT SELECT "BatchId" FROM "Recipients");
  END;"""

# Select Procedure Queries
    # IspDb Select Procedure Queries
    elif name == 'createGetServers':
        query = """\
CREATE PROCEDURE "GetServers"(IN EMAIL VARCHAR(320),
                              IN DOMAIN VARCHAR(255),
                              IN SMTP VARCHAR(4),
                              IN IMAP VARCHAR(4),
                              OUT THREAD VARCHAR(100),
                              OUT SMTPSERVER VARCHAR(255),
                              OUT IMAPSERVER VARCHAR(255),
                              OUT SMTPPORT SMALLINT,
                              OUT IMAPPORT SMALLINT,
                              OUT SMTPLOGIN VARCHAR(100),
                              OUT IMAPLOGIN VARCHAR(100),
                              OUT SMTPPWD VARCHAR(100),
                              OUT IMAPPDW VARCHAR(100))
  SPECIFIC "GetServers_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE RSLT CURSOR WITH RETURN FOR
      SELECT "Servers"."Service","Servers"."Server", "Servers"."Port", "Servers"."Connection",
      "Servers"."Authentication", "Servers"."LoginMode" FROM "Servers"
      JOIN "Providers" ON "Servers"."Provider"="Providers"."Provider"
      LEFT JOIN "Domains" ON "Providers"."Provider"="Domains"."Provider"
      WHERE "Providers"."Provider"=DOMAIN OR "Domains"."Domain"=DOMAIN
      FOR READ ONLY;
    SET (THREAD,SMTPSERVER,IMAPSERVER,SMTPPORT,IMAPPORT,SMTPLOGIN,IMAPLOGIN,SMTPPWD,IMAPPDW) =
      (SELECT U."ThreadId",S1."Server",S2."Server",S1."Port",S2."Port",S1."LoginName",S2."LoginName",S1."Password",S2."Password"
        FROM "Users" AS U
        JOIN "Services" AS S1 ON U."User"=S1."User" AND S1."Service"=SMTP
        LEFT JOIN "Services" AS S2 ON U."User"=S2."User" AND S2."Service"=IMAP
        WHERE U."User"=EMAIL);
    OPEN RSLT;
  END;"""

    # MailSpooler Select Procedure Queries
    elif name == 'createGetRecipient':
        query = """\
CREATE PROCEDURE "GetRecipient"(IN JOBID INTEGER)
  SPECIFIC "GetRecipient_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE RSLT CURSOR WITH RETURN FOR
      SELECT "Recipient", "Index", "BatchId" From "Recipients"
      WHERE "JobId"=JOBID
      FOR READ ONLY;
    OPEN RSLT;
  END;"""

    elif name == 'createGetMailer':
        query = """\
CREATE PROCEDURE "GetMailer"(IN BATCHID INTEGER,
                             IN TIMEOUT INTEGER,
                             IN SMTP VARCHAR(4),
                             IN IMAP VARCHAR(4))
  SPECIFIC "GetMailer_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE RSLT CURSOR WITH RETURN FOR
      SELECT S."Sender",S."Subject",S."Document",S."DataSource",
      S."Query",S."Table",S."Identifier",S."Bookmark",S."ThreadId",
      CASE WHEN S."DataSource" IS NULL THEN FALSE ELSE TRUE END AS "Merge",
      S1."Server" AS "SMTPServerName",S1."Port" AS "SMTPPort",
      S1."LoginName" AS "SMTPLogin",S1."Password" AS "SMTPPassword",
      C1."Connection" AS "SMTPConnectionType",A1."Authentication" AS "SMTPAuthenticationType",
      TIMEOUT AS "SMTPTimeout",
      S3."Server" AS "IMAPServerName",S3."Port" AS "IMAPPort",
      S3."LoginName" AS "IMAPLogin",S3."Password" AS "IMAPPassword",
      C2."Connection" AS "IMAPConnectionType",A2."Authentication" AS "IMAPAuthenticationType",
      TIMEOUT AS "IMAPTimeout"
      FROM "Senders" AS S
      JOIN "Services" AS S1 ON S."Sender"=S1."User" AND S1."Service"=SMTP
      JOIN "Servers" AS S2 ON S1."Server"=S2."Server" AND S1."Port"=S2."Port" AND S2."Service"=SMTP
      JOIN "ConnectionType" AS C1 ON S2."Connection"=C1."Type"
      JOIN "AuthenticationType" AS A1 ON S2."Authentication"=A1."Type"
      LEFT JOIN "Services" AS S3 ON S."Sender"=S3."User" AND S3."Service"=IMAP
      LEFT JOIN "Servers" AS S4 ON S3."Server"=S4."Server" AND S3."Port"=S4."Port" AND S4."Service"=IMAP
      LEFT JOIN "ConnectionType" AS C2 ON S4."Connection"=C2."Type"
      LEFT JOIN "AuthenticationType" AS A2 ON S4."Authentication"=A2."Type"
      WHERE S."BatchId"=BATCHID
      FOR READ ONLY;
    OPEN RSLT;
  END;"""

    elif name == 'createUpdateMailer':
        query = """\
CREATE PROCEDURE "UpdateMailer"(IN BATCHID INTEGER,
                                IN THREADID VARCHAR(100))
  SPECIFIC "UpdateMailer_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    UPDATE "Senders" SET "ThreadId"=THREADID WHERE "BatchId"=BATCHID;
  END"""

    elif name == 'createGetAttachments':
        query = """\
CREATE PROCEDURE "GetAttachments"(IN BATCHID INTEGER)
  SPECIFIC "GetAttachments_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE RSLT CURSOR WITH RETURN FOR
      SELECT "Attachment" From "Attachments"
      WHERE "BatchId"=BATCHID ORDER BY "Created"
      FOR READ ONLY;
    OPEN RSLT;
  END;"""

# Insert Procedure Queries
    # SpoolerService Insert Procedure Queries
    elif name == 'createInsertJob':
        query = """\
CREATE PROCEDURE "InsertJob"(IN SENDER VARCHAR(320),
                             IN SUBJECT VARCHAR(78),
                             IN DOCUMENT VARCHAR(512),
                             IN RECIPIENTS VARCHAR(320) ARRAY,
                             IN ATTACHMENTS VARCHAR(512) ARRAY,
                             OUT BATCHID INTEGER)
  SPECIFIC "InsertJob_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    DECLARE ID INTEGER DEFAULT 0;
    DECLARE INDEX INTEGER DEFAULT 1;
    INSERT INTO "Senders" ("Sender","Subject","Document")
    VALUES (SENDER,SUBJECT,DOCUMENT);
    SET ID = IDENTITY();
    WHILE INDEX <= CARDINALITY(RECIPIENTS) DO
      INSERT INTO "Recipients" ("BatchId","Recipient")
      VALUES (ID,RECIPIENTS[INDEX]);
      SET INDEX = INDEX + 1;
    END WHILE;
    SET INDEX = 1;
    WHILE INDEX <= CARDINALITY(ATTACHMENTS) DO
      INSERT INTO "Attachments" ("BatchId","Attachment")
      VALUES (ID,ATTACHMENTS[INDEX]);
      SET INDEX = INDEX + 1;
    END WHILE;
    SET BATCHID = ID;
  END;"""

    # SpoolerService Insert Procedure Queries
    elif name == 'createInsertMergeJob':
        query = """\
CREATE PROCEDURE "InsertMergeJob"(IN SENDER VARCHAR(320),
                                  IN SUBJECT VARCHAR(78),
                                  IN DOCUMENT VARCHAR(512),
                                  IN DATASOURCE VARCHAR(512),
                                  IN QUERYNAME VARCHAR(512),
                                  IN TABLENAME VARCHAR(512),
                                  IN IDENTIFIER VARCHAR(512),
                                  IN BOOKMARK VARCHAR(512),
                                  IN RECIPIENTS VARCHAR(320) ARRAY,
                                  IN INDEXES VARCHAR(128) ARRAY,
                                  IN ATTACHMENTS VARCHAR(512) ARRAY,
                                  OUT BATCHID INTEGER)
  SPECIFIC "InsertMergeJob_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    DECLARE ID INTEGER;
    DECLARE INDEX INTEGER DEFAULT 1;
    INSERT INTO "Senders" ("Sender","Subject","Document","DataSource","Query","Table","Identifier","Bookmark")
    VALUES (SENDER,SUBJECT,DOCUMENT,DATASOURCE,QUERYNAME,TABLENAME,IDENTIFIER,BOOKMARK);
    SET ID = IDENTITY();
    WHILE INDEX <= CARDINALITY(RECIPIENTS) DO
      INSERT INTO "Recipients" ("BatchId","Recipient","Index")
      VALUES (ID,RECIPIENTS[INDEX],INDEXES[INDEX]);
      SET INDEX = INDEX + 1;
    END WHILE;
    SET INDEX = 1;
    WHILE INDEX <= CARDINALITY(ATTACHMENTS) DO
      INSERT INTO "Attachments" ("BatchId","Attachment")
      VALUES (ID,ATTACHMENTS[INDEX]);
      SET INDEX = INDEX + 1;
    END WHILE;
    SET BATCHID = ID;
  END;"""

# Update Procedure Queries
    # IspDb Update Procedure Queries
    elif name == 'createUpdateServer':
        query = """\
CREATE PROCEDURE "UpdateServer"(IN SERVICE VARCHAR(4),
                                IN SERVER1 VARCHAR(255),
                                IN PORT1 SMALLINT,
                                IN SERVER2 VARCHAR(255),
                                IN PORT2 SMALLINT,
                                IN CONNECTION TINYINT,
                                IN AUTHENTICATION TINYINT)
  SPECIFIC "UpdateServer_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    UPDATE "Servers" SET "Server"=SERVER2,"Port"=PORT2,
      "Connection"=CONNECTION,"Authentication"=AUTHENTICATION,
      "Modified"=DEFAULT
      WHERE "Service"=SERVICE AND "Server"=SERVER1 AND "Port"=PORT1;
  END"""

# Merge Procedure Queries
    # IspDb Merge Procedure Queries
    elif name == 'createMergeProvider':
        query = """\
CREATE PROCEDURE "MergeProvider"(IN PROVIDER VARCHAR(100),
                                 IN DISPLAYNAME  VARCHAR(100),
                                 IN DISPLAYSHORTNAME VARCHAR(100))
  SPECIFIC "MergeProvider_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    MERGE INTO "Providers" USING (VALUES(PROVIDER,DISPLAYNAME,DISPLAYSHORTNAME))
      AS vals(x,y,z) ON "Provider"=vals.x
        WHEN MATCHED THEN UPDATE
          SET "DisplayName"=vals.y,"DisplayShortName"=vals.z,"Modified"=DEFAULT
        WHEN NOT MATCHED THEN INSERT
          ("Provider","DisplayName","DisplayShortName")
          VALUES vals.x,vals.y,vals.z;
  END"""

    # IspDb Merge Procedure Queries
    elif name == 'createMergeDomain':
        query = """\
CREATE PROCEDURE "MergeDomain"(IN PROVIDER VARCHAR(100),
                               IN DOMAIN VARCHAR(255))
  SPECIFIC "MergeDomain_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    MERGE INTO "Domains" USING (VALUES(DOMAIN,PROVIDER))
      AS vals(y,z) ON "Domain"=vals.y
        WHEN MATCHED THEN UPDATE
          SET "Provider"=vals.z,"Modified"=DEFAULT
        WHEN NOT MATCHED THEN INSERT
          ("Domain","Provider")
          VALUES vals.y,vals.z;
  END"""

    # IspDb Merge Procedure Queries
    elif name == 'createMergeServer':
        query = """\
CREATE PROCEDURE "MergeServer"(IN PROVIDER VARCHAR(100),
                               IN SERVICE VARCHAR(4),
                               IN SERVER VARCHAR(255),
                               IN PORT SMALLINT,
                               IN CONNECTION TINYINT,
                               IN AUTHENTICATION TINYINT,
                               IN LOGIN TINYINT)
  SPECIFIC "MergeServer_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    MERGE INTO "Servers" USING (VALUES(SERVICE,SERVER,PORT,PROVIDER,CONNECTION,AUTHENTICATION,LOGIN))
      AS vals(t,u,v,w,x,y,z) ON "Service"=vals.t AND "Server"=vals.u AND "Port"=vals.v
        WHEN MATCHED THEN UPDATE
          SET "Provider"=vals.w,"Connection"=vals.x,"Authentication"=vals.y,"LoginMode"=vals.z,"Modified"=DEFAULT
        WHEN NOT MATCHED THEN INSERT
          ("Service","Server","Port","Provider","Connection","Authentication","LoginMode")
          VALUES vals.t,vals.u,vals.v,vals.w,vals.x,vals.y,vals.z;
  END"""

    # IspDb Merge Procedure Queries
    elif name == 'createMergeUser':
        query = """\
CREATE PROCEDURE "MergeUser"(IN EMAIL VARCHAR(320),
                             IN THREAD VARCHAR(100),
                             IN SMTP VARCHAR(4),
                             IN SMTPSERVER VARCHAR(255),
                             IN SMTPPORT SMALLINT,
                             IN SMTPLOGIN VARCHAR(100),
                             IN SMTPPWD VARCHAR(100),
                             IN IMAP VARCHAR(4),
                             IN IMAPSERVER VARCHAR(255),
                             IN IMAPPORT SMALLINT,
                             IN IMAPLOGIN VARCHAR(100),
                             IN IMAPPWD VARCHAR(100))
  SPECIFIC "MergeUser_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    MERGE INTO "Users" USING (VALUES(EMAIL,THREAD))
      AS vals(y,z) ON "User"=vals.y
        WHEN MATCHED THEN UPDATE
          SET "ThreadId"=vals.z,"Modified"=DEFAULT
        WHEN NOT MATCHED THEN INSERT
          ("User","ThreadId")
          VALUES vals.y,vals.z;
    MERGE INTO "Services" USING (VALUES(EMAIL,SMTP,SMTPSERVER,SMTPPORT,SMTPLOGIN,SMTPPWD))
      AS vals(u,v,w,x,y,z) ON "User"=vals.u AND "Service"=vals.v
        WHEN MATCHED THEN UPDATE
          SET "Server"=vals.w,"Port"=vals.x,"LoginName"=vals.y,"Password"=vals.z,"Modified"=DEFAULT
        WHEN NOT MATCHED THEN INSERT
          ("User","Service","Server","Port","LoginName","Password")
          VALUES vals.u,vals.v,vals.w,vals.x,vals.y,vals.z;
    IF IMAP IS NOT NULL THEN
      MERGE INTO "Services" USING (VALUES(EMAIL,IMAP,IMAPSERVER,IMAPPORT,IMAPLOGIN,IMAPPWD))
        AS vals(u,v,w,x,y,z) ON "User"=vals.u AND "Service"=vals.v
          WHEN MATCHED THEN UPDATE
            SET "Server"=vals.w,"Port"=vals.x,"LoginName"=vals.y,"Password"=vals.z,"Modified"=DEFAULT
          WHEN NOT MATCHED THEN INSERT
            ("User","Service","Server","Port","LoginName","Password")
            VALUES vals.u,vals.v,vals.w,vals.x,vals.y,vals.z;
    END IF;
  END"""

# Call Procedure Query
    elif name == 'getServers':
        query = 'CALL "GetServers"(?,?,?,?,?,?,?,?,?,?,?,?,?)'
    elif name == 'getRecipient':
        query = 'CALL "GetRecipient"(?)'
    elif name == 'getMailer':
        query = 'CALL "GetMailer"(?,?,?,?)'
    elif name == 'getAttachments':
        query = 'CALL "GetAttachments"(?)'
    elif name == 'deleteJobs':
        query = 'CALL "DeleteJobs"(?)'
    elif name == 'insertJob':
        query = 'CALL "InsertJob"(?,?,?,?,?,?)'
    elif name == 'insertMergeJob':
        query = 'CALL "InsertMergeJob"(?,?,?,?,?,?,?,?,?,?,?,?)'
    elif name == 'updateServer':
        query = 'CALL "UpdateServer"(?,?,?,?,?,?,?)'
    elif name == 'mergeProvider':
        query = 'CALL "MergeProvider"(?,?,?)'
    elif name == 'mergeDomain':
        query = 'CALL "MergeDomain"(?,?)'
    elif name == 'mergeServer':
        query = 'CALL "MergeServer"(?,?,?,?,?,?,?)'
    elif name == 'mergeUser':
        query = 'CALL "MergeUser"(?,?,?,?,?,?,?,?,?,?,?,?)'
    elif name == 'updateMailer':
        query = 'CALL "UpdateMailer"(?,?)'

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
