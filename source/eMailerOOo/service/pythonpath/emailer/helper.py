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
import uno

from com.sun.star.uno import Exception as UnoException

from com.sun.star.mail import MailSpoolerException

from .unotool import checkVersion
from .unotool import createService
from .unotool import getExtensionVersion
from .unotool import getFileSequence
from .unotool import getMessageBox
from .unotool import getPathSettings
from .unotool import getPropertyValueSet
from .unotool import getSimpleFile
from .unotool import getToolKit
from .unotool import getUrl
from .unotool import hasInterface

from .oauth20 import getOAuth2Version
from .oauth20 import g_extension as g_oauth2ext
from .oauth20 import g_version as g_oauth2ver

from .jdbcdriver import g_extension as g_jdbcext
from .jdbcdriver import g_identifier as g_jdbcid
from .jdbcdriver import g_version as g_jdbcver

from .dbconfig import g_version

from .configuration import g_extension

import base64
import traceback


def checkOAuth2(ctx, source, logger, warn=False):
    oauth2 = getOAuth2Version(ctx)
    if oauth2 is None:
        title, msg = _getExceptionMessage(logger, 501, g_oauth2ext, g_oauth2ext, g_extension)
        if warn:
            _showWarning(ctx, title, msg)
        raise UnoException(msg, source)
    if not checkVersion(oauth2, g_oauth2ver):
        title, msg = _getExceptionMessage(logger, 503, g_oauth2ext, oauth2, g_oauth2ext, g_oauth2ver)
        if warn:
            _showWarning(ctx, title, msg)
        raise UnoException(msg, source)

def checkConfiguration(ctx, source, logger, warn=False):
    checkOAuth2(ctx, source, logger, warn)
    driver = getExtensionVersion(ctx, g_jdbcid)
    if driver is None:
        title, msg = _getExceptionMessage(logger, 501, g_jdbcext, g_jdbcext, g_extension)
        if warn:
            _showWarning(ctx, title, msg)
        raise UnoException(msg, source)
    if not checkVersion(driver, g_jdbcver):
        title, msg = _getExceptionMessage(logger, 503, g_jdbcext, driver, g_jdbcext, g_jdbcver)
        if warn:
            _showWarning(ctx, title, msg)
        raise UnoException(msg, source)

def checkConnection(ctx, source, connection, logger, new, warn=False):
    version = connection.getMetaData().getDriverVersion()
    if not checkVersion(version, g_version):
        connection.close()
        title, msg = _getExceptionMessage(logger, 511, g_jdbcext, version, g_version)
        if warn:
            _showWarning(ctx, title, msg)
        raise UnoException(msg, source)
    service = 'com.sun.star.sdb.Connection'
    interface = 'com.sun.star.sdbcx.XGroupsSupplier'
    if new and not _checkConnection(connection, service, interface):
        connection.close()
        title, msg = _getExceptionMessage(logger, 513, g_jdbcext, service, interface)
        if warn:
            _showWarning(ctx, title, msg)
        raise UnoException(msg, source)

def getDataBaseContext(ctx, source, name, resolver, code, *format):
    dbcontext = createService(ctx, 'com.sun.star.sdb.DatabaseContext')
    if not dbcontext.hasByName(name):
        msg = resolver(code, name, *format)
        raise MailSpoolerException(msg, source, ())
    location = dbcontext.getDatabaseLocation(name)
    # FIXME: The location can be an embedded database in odt file
    # FIXME: ie: vnd.sun.star.pkg://file://path/document.odt/EmbeddedDatabase
    # FIXME: eMailerOOo cannot work with such a datasource
    if not location.startswith('file://'):
        msg = resolver(code + 1, location, *format)
        raise MailSpoolerException(msg, source, ())
    # We need to check if the registered datasource has an existing odb file
    if not getSimpleFile(ctx).exists(location):
        msg = resolver(code + 2, location, *format)
        raise MailSpoolerException(msg, source, ())
    return dbcontext, location

def getMessageImage(ctx, url):
    lenght, sequence = getFileSequence(ctx, url)
    img = base64.b64encode(sequence.value).decode('utf-8')
    return img

def getDocumentFilter(extension, format):
    ext = extension.lower()
    if ext in ('odt', 'ott', 'odm', 'doc', 'dot'):
        filters = {'pdf': 'writer_pdf_Export', 'html': 'XHTML Writer File'}
    elif ext in ('ods', 'ots', 'xls', 'xlt'):
        filters = {'pdf': 'calc_pdf_Export', 'html': 'XHTML Calc File'}
    elif ext in ('odg', 'otg'):
        filters = {'pdf': 'draw_pdf_Export', 'html': 'draw_html_Export'}
    elif ext in ('odp', 'otp', 'ppt', 'pot'):
        filters = {'pdf': 'impress_pdf_Export', 'html': 'impress_html_Export'}
    else:
        filters = {}
    filter = filters.get(format, None)
    return filter

def getMailService(ctx, stype):
    service = 'com.sun.star.mail.MailServiceProvider'
    mailtype = uno.Enum('com.sun.star.mail.MailServiceType', stype)
    return createService(ctx, service).create(mailtype)

def getMailUser(ctx, sender):
    service = 'com.sun.star.mail.MailUser'
    user = createService(ctx, service, sender)
    return user

def getMailSpooler(ctx):
    service = 'com.sun.star.mail.MailSpooler'
    spooler = createService(ctx, service)
    return spooler

def getMailMessage(ctx, sender, recipient, subject, body):
    service = 'com.sun.star.mail.MailMessage'
    mail = createService(ctx, service, recipient, sender, subject, body)
    return mail

def getTransferable(ctx):
    service = 'com.sun.star.datatransfer.TransferableFactory'
    transferable = createService(ctx, service)
    return transferable

def getNamedExtension(name):
    part1, dot, part2 = name.rpartition('.')
    if dot:
        name, extension = part1, part2
    else:
        name, extension = part2, None
    return name, extension

def saveDocumentAs(ctx, document, format):
    url = None
    name, extension = getNamedExtension(document.Title)
    if extension is None:
        extension = _getDocumentExtension(document)
    filter = getDocumentFilter(extension, format)
    if filter is not None:
        temp = getPathSettings(ctx).Temp
        url = '%s/%s.%s' % (temp, name, format)
        descriptor = getPropertyValueSet({'FilterName': filter, 'Overwrite': True})
        document.storeToURL(url, descriptor)
        url = getUrl(ctx, url)
        if url is not None:
            url = url.Main
    return url

def saveTempDocument(document, url, name, format=None):
    descriptor = {'Overwrite': True}
    title, extension = getNamedExtension(name)
    if extension is None:
        extension = _getDocumentExtension(document)
    filter = getDocumentFilter(extension, format)
    if filter is not None:
        descriptor['FilterName'] = filter
    document.storeToURL(url, getPropertyValueSet(descriptor))
    return name if format is None else '%s.%s' % (title, format)

def saveDocumentTmp(ctx, document, format=None):
    url = None
    descriptor = {'Overwrite': True}
    name, extension = getNamedExtension(document.Title)
    if extension is None:
        extension = _getDocumentExtension(document)
    filter = getDocumentFilter(extension, format)
    if filter is not None:
        descriptor['FilterName'] = filter
    temp = getPathSettings(ctx).Temp
    if format is None:
        url = '%s/%s' % (temp, document.Title)
    else:
        url = '%s/%s.%s' % (temp, name, format)
    document.storeToURL(url, getPropertyValueSet(descriptor))
    url = getUrl(ctx, url)
    return url

def getTransferableMimeValues(detection, descriptor, uiname, mimetype, deep=True):
    itype, descriptor = detection.queryTypeByDescriptor(getPropertyValueSet(descriptor), deep)
    if detection.hasByName(itype):
        for t in detection.getByName(itype):
            if t.Name == 'UIName':
                uiname = t.Value
            elif t.Name == 'MediaType':
                mimetype = t.Value
    return uiname, mimetype

def _getDocumentExtension(document):
    identifier = document.getIdentifier()
    if identifier == 'com.sun.star.text.TextDocument':
        extension = 'odt'
    elif identifier == 'com.sun.star.sheet.SpreadsheetDocument':
        extension = 'ods'
    elif identifier == 'com.sun.star.drawing.DrawingDocument':
        extension = 'odg'
    elif identifier == 'com.sun.star.presentation.PresentationDocument':
        extension = 'odp'
    else:
        extension = None
    return extension

def _checkJdbc(ctx, method):
    driver = getExtensionVersion(ctx, g_jdbcid)
    if driver is None:
        return method, 501, g_jdbcext, g_extension
    elif not checkVersion(driver, g_jdbcver):
        return method, 503, driver, g_jdbcext, g_jdbcver
    return None

def _getExceptionMessage(logger, code, extension, *args):
    title = logger.resolveString(code, extension)
    message = logger.resolveString(code + 1, *args)
    return title, message

def _showWarning(ctx, title, msg):
    toolkit = getToolKit(ctx)
    peer = toolkit.getActiveTopWindow()
    box = uno.Enum('com.sun.star.awt.MessageBoxType', 'ERRORBOX')
    msgbox = getMessageBox(toolkit, peer, box, 1, title, msg)
    msgbox.execute()
    msgbox.dispose()

def _checkConnection(connection, service, interface):
    return connection.supportsService(service) and hasInterface(connection, interface)

