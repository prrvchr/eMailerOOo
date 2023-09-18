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

from .datasource import DataSource

from .unotool import checkVersion
from .unotool import createService
from .unotool import getExtensionVersion
from .unotool import getFileSequence
from .unotool import getPathSettings
from .unotool import getPropertyValueSet
from .unotool import getTypeDetection
from .unotool import getUrl

from .oauth2 import getOAuth2Version
from .oauth2 import g_extension as oauth2ext
from .oauth2 import g_version as oauth2ver

from .jdbcdriver import g_extension as jdbcext
from .jdbcdriver import g_identifier as jdbcid
from .jdbcdriver import g_version as jdbcver

from .dbconfig import g_version

from .configuration import g_extension

import base64
import traceback


def checkSetup(ctx, method, sep, callback):
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
        datasource = DataSource(ctx)
        if not datasource.DataBase.canConnect():
            callback(method, 505, datasource.DataBase.Url, sep, datasource.DataBase.Error)
        elif not datasource.DataBase.isUptoDate():
            callback(method, 507, datasource.DataBase.Version, sep, g_version)
        else:
            return datasource
    return None

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

def getMail(ctx, sender, recipient, subject, body):
    service = 'com.sun.star.mail.MailMessage2'
    mail = createService(ctx, service)
    mail.create(recipient, sender, subject, body)
    return mail

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

def getAttachmentType(ctx, url):
    service = 'com.sun.star.document.TypeDetection'
    detection = createService(ctx, service)
    type = detection.queryTypeByURL(url)
    if detection.hasByName(type):
        types = detection.getByName(type)
        for type in types:
            if type.Name == 'MediaType':
                return type.Value
    return 'application/octet-stream'

def getUrlMimeType(detection, url):
    mimetype = 'application/octet-stream'
    name = detection.queryTypeByURL(url)
    if detection.hasByName(name):
        types = detection.getByName(name)
        for t in types:
            print("mailertool.getUrlMimeType() %s - %s" % (t.Name, t.Value))
            if t.Name == 'MediaType':
                mimetype = t.Value
    return mimetype

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
