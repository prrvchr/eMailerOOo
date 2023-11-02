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
import uno

from com.sun.star.sdbc import SQLException

from com.sun.star.mail import MailSpoolerException

from .database import DataBase

from .datasource import DataSource

from .unotool import checkVersion
from .unotool import createService
from .unotool import getExtensionVersion
from .unotool import getFileSequence
from .unotool import getPathSettings
from .unotool import getPropertyValueSet
from .unotool import getSimpleFile
from .unotool import getTypeDetection
from .unotool import getUrl

from .dbtool import getConnectionUrl

from .oauth2 import getOAuth2Version
from .oauth2 import g_extension as oauth2ext
from .oauth2 import g_version as oauth2ver

from .jdbcdriver import g_extension as jdbcext
from .jdbcdriver import g_identifier as jdbcid
from .jdbcdriver import g_version as jdbcver

from .dbconfig import g_folder
from .dbconfig import g_version

from .configuration import g_extension
from .configuration import g_basename

import base64
import traceback


def getDataSource(ctx, method, sep, callback):
    oauth2 = getOAuth2Version(ctx)
    driver = getExtensionVersion(ctx, jdbcid)
    if oauth2 is None:
        callback(method, 501, oauth2ext, sep, g_extension)
    elif not checkVersion(oauth2, oauth2ver):
        callback(method, 503, oauth2, oauth2ext, sep, oauth2ver)
    elif driver is None:
        callback(method, 501, jdbcext, sep, g_extension)
    elif not checkVersion(driver, jdbcver):
        callback(method, 503, driver, jdbcext, sep, jdbcver)
    else:
        path = g_folder + '/' + g_basename
        url = getConnectionUrl(ctx, path)
        try:
            database = DataBase(ctx, url)
        except SQLException as e:
            callback(method, 505, url, sep, e.Message)
        else:
            if not database.isUptoDate():
                callback(method, 507, database.Version, sep, g_version)
            else:
                return DataSource(ctx, database)
    return None

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
        msg = resolver(code +1, location, *format)
        raise MailSpoolerException(msg, source, ())
    # We need to check if the registered datasource has an existing odb file
    if not getSimpleFile(ctx).exists(location):
        msg = resolver(code +2, location, *format)
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

def getMailService(ctx, service):
    name = 'com.sun.star.mail.MailServiceProvider'
    stype = uno.Enum('com.sun.star.mail.MailServiceType', service)
    return createService(ctx, name).create(stype)

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
    print("mailtool.saveTempDocument() %s - %s" % (format, filter))
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
    print("mailtool.saveDocumentTmp() %s - %s" % (format, filter))
    temp = getPathSettings(ctx).Temp
    if format is None:
        url = '%s/%s' % (temp, document.Title)
    else:
        url = '%s/%s.%s' % (temp, name, format)
    document.storeToURL(url, getPropertyValueSet(descriptor))
    url = getUrl(ctx, url)
    return url

def getTransferableMimeValues(ctx, descriptor, uiname, deep=True):
    mimetype = 'application/octet-stream'
    service = 'com.sun.star.document.TypeDetection'
    detection = createService(ctx, service)
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
