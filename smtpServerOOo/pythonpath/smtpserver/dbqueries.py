#!
# -*- coding: utf_8 -*-

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .logger import logMessage
from .logger import getMessage
g_message = 'dbqueries'


def getSqlQuery(ctx, name, format=None):

# Select Queries
    if name == 'getTablesName':
        query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.SYSTEM_TABLES WHERE TABLE_TYPE='TABLE'"

# Get DataBase Version Query
    elif name == 'getVersion':
        query = 'SELECT DISTINCT DATABASE_VERSION() AS "HSQL Version" FROM INFORMATION_SCHEMA.SYSTEM_TABLES'

# ShutDown Queries
    elif name == 'shutdown':
        query = 'SHUTDOWN;'
    elif name == 'shutdownCompact':
        query = 'SHUTDOWN COMPACT;'

# Queries don't exist!!!
    else:
        query = None
        msg = getMessage(ctx, g_message, 101, name)
        logMessage(ctx, SEVERE, msg, 'dbqueries', 'getSqlQuery()')
    return query
