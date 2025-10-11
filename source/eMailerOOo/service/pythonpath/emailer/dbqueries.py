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
    DELETE FROM "Addresses" WHERE "BatchId" IN (SELECT "BatchId" FROM "Addresses"
    EXCEPT SELECT "BatchId" FROM "Recipients");
    DELETE FROM "Attachments" WHERE "BatchId" IN (SELECT "BatchId" FROM "Attachments"
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
                                 IN RECIPIENT VARCHAR(320),
                                 IN THREADID VARCHAR(256),
                                 IN MESSAGEID VARCHAR(256),
                                 IN FOREIGNDID VARCHAR(256),
                                 IN STATE INTEGER)
  SPECIFIC "UpdateSpooler_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    UPDATE "Senders" SET "ThreadId" = THREADID WHERE "BatchId" = BATCHID;
    UPDATE "Recipients" SET "State" = STATE, "Recipient" = RECIPIENT, "MessageId" = MESSAGEID, "ForeignId" = FOREIGNDID, "Modified" = DEFAULT WHERE "JobId" = JOBID;
  END"""

    elif name == 'createUpdateJobState':
        query = """\
CREATE PROCEDURE "UpdateJobState"(IN JOBID INTEGER,
                                  IN STATE INTEGER)
  SPECIFIC "UpdateJobState_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    UPDATE "Recipients" SET "State"=STATE, "Modified"=DEFAULT WHERE "JobId"=JOBID;
  END"""

    elif name == 'createUpdateJobsState':
        query = """\
CREATE PROCEDURE "UpdateJobsState"(IN JOBIDS INTEGER ARRAY,
                                   IN STATE INTEGER)
  SPECIFIC "UpdateJobsState_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    UPDATE "Recipients" SET "State"=STATE, "Modified"=DEFAULT WHERE "JobId" IN (UNNEST (JOBIDS));
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
    INSERT INTO "Senders" ("Sender", "Subject", "Document")
    VALUES (SENDER, SUBJECT, DOCUMENT);
    SET ID = IDENTITY();
    WHILE INDEX <= CARDINALITY(RECIPIENTS) DO
      INSERT INTO "Recipients" ("BatchId", "Recipient")
      VALUES (ID, RECIPIENTS[INDEX]);
      SET INDEX = INDEX + 1;
    END WHILE;
    SET INDEX = 1;
    WHILE INDEX <= CARDINALITY(ATTACHMENTS) DO
      INSERT INTO "Attachments" ("BatchId", "Attachment")
      VALUES (ID, ATTACHMENTS[INDEX]);
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
                                  IN PREDICATES VARCHAR(256) ARRAY,
                                  IN ADDRESSES VARCHAR(128) ARRAY,
                                  IN IDENTIFIERS VARCHAR(128) ARRAY,
                                  IN ATTACHMENTS VARCHAR(512) ARRAY,
                                  OUT BATCHID INTEGER)
  SPECIFIC "InsertMergeJob_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    DECLARE BATCH INTEGER;
    DECLARE INDEX INTEGER DEFAULT 1;
    INSERT INTO "Senders" ("Sender", "Subject", "Document", "DataSource", "Query", "Table")
    VALUES (SENDER, SUBJECT, DOCUMENT, DATASOURCE, QUERYNAME, TABLENAME);
    SET BATCH = IDENTITY();
    WHILE INDEX <= CARDINALITY(RECIPIENTS) DO
      INSERT INTO "Recipients" ("BatchId", "Recipient", "Filter", "Predicate")
      VALUES (BATCH, RECIPIENTS[INDEX], FILTERS[INDEX], PREDICATES[INDEX]);
      SET INDEX = INDEX + 1;
    END WHILE;
    SET INDEX = 1;
    WHILE INDEX <= CARDINALITY(ADDRESSES) DO
      INSERT INTO "Addresses" ("BatchId", "Address")
      VALUES (BATCH, ADDRESSES[INDEX]);
      SET INDEX = INDEX + 1;
    END WHILE;
    SET INDEX = 1;
    WHILE INDEX <= CARDINALITY(IDENTIFIERS) DO
      INSERT INTO "Identifiers" ("BatchId", "Identifier")
      VALUES (BATCH, IDENTIFIERS[INDEX]);
      SET INDEX = INDEX + 1;
    END WHILE;
    SET INDEX = 1;
    WHILE INDEX <= CARDINALITY(ATTACHMENTS) DO
      INSERT INTO "Attachments" ("BatchId", "Attachment")
      VALUES (BATCH, ATTACHMENTS[INDEX]);
      SET INDEX = INDEX + 1;
    END WHILE;
    SET BATCHID = BATCH;
  END;"""

    elif name == 'createResubmitJobs':
        query = """\
CREATE PROCEDURE "ResubmitJobs"(IN JOBIDS INTEGER ARRAY,
                                OUT IDS INTEGER ARRAY)
  SPECIFIC "ResubmitJobs_1"
  MODIFIES SQL DATA
  BEGIN ATOMIC
    DECLARE BATCHIDS INTEGER ARRAY DEFAULT ARRAY[];
    DECLARE RECIPIENTS VARCHAR(320) ARRAY;
    DECLARE FILTERS, PREDICATES VARCHAR(256) ARRAY;
    DECLARE ADDRESSES, IDENTIFIERS VARCHAR(128) ARRAY;
    DECLARE ATTACHMENTS VARCHAR(512) ARRAY;
    DECLARE BATCHID INTEGER;
    DECLARE INDEX INTEGER DEFAULT 1;
    DECLARE SENDER VARCHAR(320);
    DECLARE SUBJECT VARCHAR(78);
    DECLARE DOCUMENT, DATASOURCE, QUERYNAME, TABLENAME VARCHAR(512);
    DECLARE ISMERGE BOOLEAN;
    FOR_EACH_BATCH: FOR SELECT DISTINCT "BatchId" AS BATCH FROM "Recipients" WHERE "JobId" IN (UNNEST(JOBIDS)) ORDER BY "BatchId" DO
      SELECT "Sender", "Subject", "Document", "DataSource", "Query", "Table", "Merge", "Recipients", "Filters", "Predicates"
      INTO SENDER, SUBJECT, DOCUMENT, DATASOURCE, QUERYNAME, TABLENAME, ISMERGE, RECIPIENTS, FILTERS, PREDICATES
      FROM TABLE("GetBatchJobs"(JOBIDS)) WHERE "Batch" = BATCH;
      SELECT ARRAY_AGG("Attachment" ORDER BY "Created") INTO ATTACHMENTS FROM "Attachments" WHERE "BatchId" = BATCH;
      IF ISMERGE THEN
        SELECT ARRAY_AGG("Address" ORDER BY "Created") INTO ADDRESSES FROM "Addresses" WHERE "BatchId" = BATCH;
        SELECT ARRAY_AGG("Identifier" ORDER BY "Created") INTO IDENTIFIERS FROM "Identifiers" WHERE "BatchId" = BATCH;
        CALL "InsertMergeJob"(SENDER, SUBJECT, DOCUMENT, DATASOURCE, QUERYNAME, TABLENAME,
                              RECIPIENTS, FILTERS, PREDICATES, ADDRESSES, IDENTIFIERS, ATTACHMENTS, BATCHID);
      ELSE
        CALL "InsertJob"(SENDER, SUBJECT, DOCUMENT, RECIPIENTS, ATTACHMENTS, BATCHID);
      END IF;
      SET BATCHIDS[INDEX] = BATCHID;
      SET INDEX = INDEX + 1;
    END FOR FOR_EACH_BATCH;
    SET IDS = BATCHIDS;
  END;"""

    elif name == 'createGetSendJobs':
        query = """\
CREATE PROCEDURE "GetSendJobs"()
  SPECIFIC "GetSendJobs_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE RSLT CURSOR WITH RETURN FOR
      SELECT B."Batch", B."Sender", B."Subject", B."Document", B."DataSource", B."Query", B."Table", B."Merge",
      B."Jobs", B."Recipients", B."Filters", B."Predicates",
      ARRAY(SELECT "Address" FROM "Addresses" WHERE "BatchId" = B."Batch" ORDER BY "Created") AS "Addresses",
      ARRAY(SELECT "Identifier" FROM "Identifiers" WHERE "BatchId" = B."Batch" ORDER BY "Created") AS "Identifiers",
      ARRAY(SELECT "Attachment" FROM "Attachments" WHERE "BatchId" = B."Batch" ORDER BY "Created") AS "Attachments"
      FROM TABLE("GetBatchState"(0)) AS B
      ORDER BY B."Batch"
      FOR READ ONLY;
    OPEN RSLT;
  END;"""

    elif name == 'createGetJobs':
        query = """\
CREATE PROCEDURE "GetJobs"(IN JOBIDS INTEGER ARRAY)
  SPECIFIC "GetJobs_1"
  READS SQL DATA
  DYNAMIC RESULT SETS 1
  BEGIN ATOMIC
    DECLARE RSLT CURSOR WITH RETURN FOR
      SELECT B."Batch", B."Sender", B."Subject", B."Document", B."DataSource", B."Query", B."Table", B."Merge",
      B."Jobs", B."Recipients", B."Filters", B."Predicates",
      ARRAY(SELECT "Address" FROM "Addresses" WHERE "BatchId" = B."Batch" ORDER BY "Created") AS "Addresses",
      ARRAY(SELECT "Identifier" FROM "Identifiers" WHERE "BatchId" = B."Batch" ORDER BY "Created") AS "Identifiers",
      ARRAY(SELECT "Attachment" FROM "Attachments" WHERE "BatchId" = B."Batch" ORDER BY "Created") AS "Attachments"
      FROM TABLE("GetBatchJobs"(JOBIDS)) AS B
      ORDER BY B."Batch"
      FOR READ ONLY;
    OPEN RSLT;
  END;"""


    # Internal Function used in Procedure
    elif name == 'createGetBatchJobs':
        query = """\
CREATE FUNCTION "GetBatchJobs"(IN JOBIDS INTEGER ARRAY)
  RETURNS TABLE
    ("Batch" INTEGER, "Sender" VARCHAR(320), "Subject" VARCHAR(78), "Document" VARCHAR(512),
     "DataSource" VARCHAR(512), "Query" VARCHAR(512), "Table" VARCHAR(512), "Merge" BOOLEAN,
     "Jobs" INTEGER ARRAY, "Recipients" VARCHAR(320) ARRAY, "Filters" VARCHAR(256) ARRAY,
     "Predicates" VARCHAR(256) ARRAY)
  SPECIFIC "GetBatchJobs_1"
  READS SQL DATA
  BEGIN ATOMIC
    RETURN TABLE
      (SELECT S."BatchId" AS "Batch", S."Sender", S."Subject", S."Document", S."DataSource",
       S."Query", S."Table", CASE WHEN S."Query" IS NULL THEN FALSE ELSE TRUE END AS "Merge",
       ARRAY_AGG(R."JobId" ORDER BY R."JobId") AS "Jobs",
       ARRAY_AGG(R."Recipient" ORDER BY R."JobId") AS "Recipients",
       ARRAY_AGG(R."Filter" ORDER BY R."JobId") AS "Filters",
       ARRAY_AGG(R."Predicate" ORDER BY R."JobId") AS "Predicates"
       FROM "Senders" AS S
       INNER JOIN "Recipients" AS R ON S."BatchId" = R."BatchId"
       WHERE R."JobId" IN (UNNEST(JOBIDS))
       GROUP BY S."BatchId", S."Sender", S."Subject", S."Document", S."DataSource", S."Query", S."Table"
       ORDER BY S."BatchId");
  END"""

    elif name == 'createGetBatchState':
        query = """\
CREATE FUNCTION "GetBatchState"(IN STATE INTEGER)
  RETURNS TABLE
    ("Batch" INTEGER, "Sender" VARCHAR(320), "Subject" VARCHAR(78), "Document" VARCHAR(512),
     "DataSource" VARCHAR(512), "Query" VARCHAR(512), "Table" VARCHAR(512), "Merge" BOOLEAN,
     "Jobs" INTEGER ARRAY, "Recipients" VARCHAR(320) ARRAY, "Filters" VARCHAR(256) ARRAY,
     "Predicates" VARCHAR(256) ARRAY)
  SPECIFIC "GetBatchState_1"
  READS SQL DATA
  BEGIN ATOMIC
    RETURN TABLE
      (SELECT S."BatchId" AS "Batch", S."Sender", S."Subject", S."Document", S."DataSource",
       S."Query", S."Table", CASE WHEN S."Query" IS NULL THEN FALSE ELSE TRUE END AS "Merge",
       ARRAY_AGG(R."JobId" ORDER BY R."JobId") AS "Jobs",
       ARRAY_AGG(R."Recipient" ORDER BY R."JobId") AS "Recipients",
       ARRAY_AGG(R."Filter" ORDER BY R."JobId") AS "Filters",
       ARRAY_AGG(R."Predicate" ORDER BY R."JobId") AS "Predicates"
       FROM "Senders" AS S
       INNER JOIN "Recipients" AS R ON S."BatchId" = R."BatchId"
       WHERE R."State" = STATE
       GROUP BY S."BatchId", S."Sender", S."Subject", S."Document", S."DataSource", S."Query", S."Table"
       ORDER BY S."BatchId");
  END"""


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
        query = 'CALL "InsertMergeJob"(?,?,?,?,?,?,?,?,?,?,?,?,?)'
    elif name == 'updateSpooler':
        query = 'CALL "UpdateSpooler"(?,?,?,?,?,?,?)'
    elif name == 'updateJobState':
        query = 'CALL "UpdateJobState"(?,?)'
    elif name == 'updateJobsState':
        query = 'CALL "UpdateJobsState"(?,?)'
    elif name == 'getSendJobs':
        query = 'CALL "GetSendJobs"()'
    elif name == 'getJobs':
        query = 'CALL "GetJobs"(?)'
    elif name == 'resubmitJobs':
        query = 'CALL "ResubmitJobs"(?,?)'

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
        logger.logprb(SEVERE, g_basename, 'getSqlQuery', 101, name)
        query = None
    return query

