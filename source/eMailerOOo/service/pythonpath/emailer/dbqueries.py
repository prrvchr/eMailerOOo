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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .logger import getLogger

from .configuration import g_errorlog
g_basename = 'dbqueries'


def getSqlQuery(ctx, name, format=None):

# DataBase creation Queries
    # Create View Queries
    if name == 'getSpoolerViewCommand':
        c1 = 'R."JobId"'
        c2 = 'S."BatchId"'
        c3 = 'R."State"'
        c4 = 'S."Subject"'
        c5 = 'S."Sender"'
        c6 = 'R."Recipient"'
        c7 = 'S."Document"'
        c8 = 'S."DataSource"'
        c9 = 'S."Query"'
        c10 = 'S."Table"'
        c11 = 'R."Filter"'
        c12 = 'R."MessageId"'
        c13 = 'R."ForeignId"'
        c14 = 'S."ThreadId"'
        c15 = 'S."Created" AS "Submit"'
        c16 = 'R."Modified" AS "Sending"'
        c = (c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11,c12,c13,c14,c15,c16)
        f1 = '%(Catalog)s.%(Schema)s."Senders" AS S'
        f2 = 'INNER JOIN %(Catalog)s.%(Schema)s."Recipients" AS R ON R."BatchId" = S."BatchId"'
        f = (f1,f2)
        p = (', '.join(c), ' '.join(f))
        command = 'SELECT %s FROM %s;' % p
        query = command % format

# Select Queries
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

    elif name == 'createUpdateSpooler':
        query = """\
CREATE PROCEDURE "UpdateSpooler"(IN BATCHID INTEGER,
                                 IN JOBID INTEGER,
                                 IN THREADID VARCHAR(256),
                                 IN MESSAGEID VARCHAR(256),
                                 IN FOREIGNDID VARCHAR(256),
                                 IN STATE INTEGER)
  SPECIFIC "UpdateSpooler_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    UPDATE "Senders" SET "ThreadId"=THREADID WHERE "BatchId"=BATCHID;
    UPDATE "Recipients" SET "State"=STATE, "MessageId"=MESSAGEID, "ForeignId"=FOREIGNDID, "Modified"=DEFAULT WHERE "JobId"=JOBID;
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
    elif name == 'updateSpooler':
        query = 'CALL "UpdateSpooler"(?,?,?,?,?,?)'

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

# Queries don't exist!!!
    else:
        logger = getLogger(ctx, g_errorlog, g_basename)
        logger.logprb(SEVERE, g_basename, 'getSqlQuery()', 101, name)
        query = None
    return query

