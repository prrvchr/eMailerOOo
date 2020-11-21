#!
# -*- coding: utf_8 -*-

from com.sun.star.logging.LogLevel import SEVERE
from com.sun.star.logging.LogLevel import WARNING
from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import CONFIG
from com.sun.star.logging.LogLevel import FINE
from com.sun.star.logging.LogLevel import FINER
from com.sun.star.logging.LogLevel import FINEST
from com.sun.star.logging.LogLevel import ALL
from com.sun.star.logging.LogLevel import OFF

from unolib import getConfiguration
from unolib import getStringResource

from .configuration import g_identifier

g_stringResource = {}
g_pathResource = 'resource'
g_logger = '%s.Logger' % g_identifier
g_loggerService = None
g_debugMode = False


def setDebugMode(mode):
    global g_debugMode
    g_debugMode = mode

def isDebugMode():
    return g_debugMode

def getMessage(ctx, fileresource, resource, format=()):
    msg = _getResource(ctx, fileresource).resolveString('%s' % resource)
    if format:
        msg = msg % format
    return msg

def logMessage(ctx, level, msg, cls=None, method=None):
    logger = _getLogger(ctx)
    if logger.isLoggable(level):
        if cls is None or method is None:
            logger.log(level, msg)
        else:
            logger.logp(level, cls, method, msg)

def clearLogger():
    global g_loggerService
    g_loggerService = None

def isLoggerEnabled(ctx):
    level = _getLoggerConfiguration(ctx).LogLevel
    enabled = _isLoggerEnabled(level)
    return enabled

def getLoggerSetting(ctx):
    configuration = _getLoggerConfiguration(ctx)
    enabled, index = _getLogIndex(configuration)
    handler = _getLogHandler(configuration)
    return enabled, index, handler

def setLoggerSetting(ctx, enabled, index, handler):
    configuration = _getLoggerConfiguration(ctx)
    _setLogIndex(configuration, enabled, index)
    _setLogHandler(configuration, handler, index)
    if configuration.hasPendingChanges():
        configuration.commitChanges()
        clearLogger()

def getLoggerUrl(ctx):
    url = '$(userurl)/$(loggername).log'
    settings = _getLoggerConfiguration(ctx).getByName('HandlerSettings')
    if settings.hasByName('FileURL'):
        url = settings.getByName('FileURL')
    service = ctx.ServiceManager.createInstance('com.sun.star.util.PathSubstitution')
    return service.substituteVariables(url.replace('$(loggername)', g_logger), True)

def _getLogger(ctx):
    if g_loggerService is None:
        _setLogger(ctx)
    return g_loggerService

def _setLogger(ctx):
    global g_loggerService
    singleton = '/singletons/com.sun.star.logging.LoggerPool'
    g_loggerService = ctx.getValueByName(singleton).getNamedLogger(g_logger)

def _getResource(ctx, fileresource):
    if fileresource not in g_stringResource:
        resource = getStringResource(ctx, g_identifier, g_pathResource, fileresource)
        g_stringResource[fileresource] = resource
    return g_stringResource[fileresource]

def _getLoggerConfiguration(ctx):
    nodepath = '/org.openoffice.Office.Logging/Settings'
    configuration = getConfiguration(ctx, nodepath, True)
    if not configuration.hasByName(g_logger):
        configuration.insertByName(g_logger, configuration.createInstance())
        configuration.commitChanges()
    nodepath += '/%s' % g_logger
    return getConfiguration(ctx, nodepath, True)

def _getLogIndex(configuration):
    index = 7
    level = configuration.LogLevel
    enabled = _isLoggerEnabled(level)
    if enabled:
        index = _getLogLevels().index(level)
    return enabled, index

def _setLogIndex(configuration, enabled, index):
    level = _getLogLevels()[index] if enabled else OFF
    if configuration.LogLevel != level:
        configuration.LogLevel = level

def _getLogHandler(configuration):
    handler = 1 if configuration.DefaultHandler != 'com.sun.star.logging.FileHandler' else 2
    return handler

def _setLogHandler(configuration, console, index):
    handler = 'com.sun.star.logging.ConsoleHandler' if console else 'com.sun.star.logging.FileHandler'
    if configuration.DefaultHandler != handler:
        configuration.DefaultHandler = handler
    settings = configuration.getByName('HandlerSettings')
    if settings.hasByName('Threshold'):
        if settings.getByName('Threshold') != index:
            settings.replaceByName('Threshold', index)
    else:
        settings.insertByName('Threshold', index)

def _getLogLevels():
    levels = (SEVERE,
              WARNING,
              INFO,
              CONFIG,
              FINE,
              FINER,
              FINEST,
              ALL)
    return levels

def _isLoggerEnabled(level):
    return level != OFF
