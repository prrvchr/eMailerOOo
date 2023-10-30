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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .logger import getLogger

from .configuration import g_errorlog
g_basename = 'dbqueries'


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
        query = 'SET TABLE "%s" SOURCE "%s"' % format

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
        c10 = '"Table"'
        c11 = '"Filter"'
        c12 = '"Submit"'
        c13 = '"Sending"'
        c = (c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11,c12,c13)
        s1 = '"Recipients"."JobId"'
        s2 = '"Senders"."BatchId"'
        s3 = '"Recipients"."State"'
        s4 = '"Senders"."Subject"'
        s5 = '"Senders"."Sender"'
        s6 = '"Recipients"."Recipient"'
        s7 = '"Senders"."Document"'
        s8 = '"Senders"."DataSource"'
        s9 = '"Senders"."Query"'
        s10 = '"Senders"."Table"'
        s11 = '"Recipients"."Filter"'
        s12 = '"Senders"."Created"'
        s13 = '"Recipients"."Modified"'
        s = (s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12,s13)
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

    # Merger Composer Select Queries
    elif name == 'getQueryCommand':
        query = 'SELECT %s.* FROM %s;' % format

    elif name == 'getRecipientQuery':
        query = 'SELECT %s AS "Recipient" FROM %s ORDER BY %s;' % format

    # Spooler Select Queries
    elif name == 'getViewQuery':
        query = 'SELECT * FROM "Spooler";'

    elif name == 'getSpoolerViewQuery':
        query = 'SELECT * FROM "View";'

    # MailSpooler Select Queries
    elif name == 'getSpoolerJobs':
        query = 'SELECT "JobId" FROM "Spooler" WHERE "State" = ? ORDER BY "JobId";'

    elif name == 'getJobState':
        query = 'SELECT "State" FROM "Recipients" WHERE "JobId" = ?;'

    elif name == 'getJobIds':
        query = 'SELECT ARRAY_AGG("JobId") FROM "Recipients" WHERE "BatchId" = ?;'

# Update Queries
    # MailSpooler Update Queries
    elif name == 'updateRecipient':
        query = 'UPDATE "Recipients" SET "State"=?, "MessageId"=?, "Modified"=? WHERE "JobId"=?;'

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
    # MailSpooler Delete Procedure Queries
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
    # MailSpooler Select Procedure Queries
    elif name == 'createGetRecipient':
        query = """\
CREATE PROCEDURE "GetRecipient"(IN JOBID INTEGER)
  SPECIFIC "GetRecipient_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE RSLT CURSOR WITH RETURN FOR
      SELECT "Recipient", "Filter", "BatchId" From "Recipients"
      WHERE "JobId"=JOBID
      FOR READ ONLY;
    OPEN RSLT;
  END;"""

    elif name == 'createGetMailer':
        query = """\
CREATE PROCEDURE "GetMailer"(IN BATCHID INTEGER)
  SPECIFIC "GetMailer_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE RSLT CURSOR WITH RETURN FOR
      SELECT "Sender", "Subject", "Document",
      "DataSource", "Query", "Table", "ThreadId",
      CASE WHEN "DataSource" IS NULL OR "Query" IS NULL OR "Table" IS NULL THEN FALSE ELSE TRUE END AS "Merge"
      FROM "Senders"
      WHERE "BatchId"=BATCHID
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
    # MailSpooler Insert Procedure Queries
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

    # MailSpooler Insert Procedure Queries
    elif name == 'createInsertMergeJob':
        query = """\
CREATE PROCEDURE "InsertMergeJob"(IN SENDER VARCHAR(320),
                                  IN SUBJECT VARCHAR(78),
                                  IN DOCUMENT VARCHAR(512),
                                  IN DATASOURCE VARCHAR(512),
                                  IN QUERYNAME VARCHAR(512),
                                  IN TABLENAME VARCHAR(512),
                                  IN RECIPIENTS VARCHAR(320) ARRAY,
                                  IN FILTERS VARCHAR(256) ARRAY,
                                  IN ATTACHMENTS VARCHAR(512) ARRAY,
                                  OUT BATCHID INTEGER)
  SPECIFIC "InsertMergeJob_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    DECLARE ID INTEGER;
    DECLARE INDEX INTEGER DEFAULT 1;
    INSERT INTO "Senders" ("Sender","Subject","Document","DataSource","Query","Table")
    VALUES (SENDER,SUBJECT,DOCUMENT,DATASOURCE,QUERYNAME,TABLENAME);
    SET ID = IDENTITY();
    WHILE INDEX <= CARDINALITY(RECIPIENTS) DO
      INSERT INTO "Recipients" ("BatchId","Recipient","Filter")
      VALUES (ID,RECIPIENTS[INDEX],FILTERS[INDEX]);
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

# Call Procedure Query
    elif name == 'getRecipient':
        query = 'CALL "GetRecipient"(?)'
    elif name == 'getMailer':
        query = 'CALL "GetMailer"(?)'
    elif name == 'getAttachments':
        query = 'CALL "GetAttachments"(?)'
    elif name == 'deleteJobs':
        query = 'CALL "DeleteJobs"(?)'
    elif name == 'insertJob':
        query = 'CALL "InsertJob"(?,?,?,?,?,?)'
    elif name == 'insertMergeJob':
        query = 'CALL "InsertMergeJob"(?,?,?,?,?,?,?,?,?,?)'
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
        logger = getLogger(ctx, g_errorlog, g_basename)
        logger.logprb(SEVERE, g_basename, 'getSqlQuery()', 101, name)
        query = None
    return query

