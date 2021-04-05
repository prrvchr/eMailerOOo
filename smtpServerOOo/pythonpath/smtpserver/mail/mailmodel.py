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

import uno
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from smtpserver import createService
from smtpserver import getMessage
from smtpserver import getPathSettings
from smtpserver import getPropertyValueSet
from smtpserver import getUrl
from smtpserver import getUrlTransformer
from smtpserver import logMessage
from smtpserver import parseUrl

from collections import OrderedDict
import validators
import json
import traceback


class MailModel(unohelper.Base):
    def __init__(self):
        raise NotImplementedError('Need to be implemented!')

    @property
    def DataSource(self):
        return self._datasource
    @property
    def Path(self):
        return self._path
    @Path.setter
    def Path(self, path):
        self._path = path

# MailModel getter methods
    def getUrl(self):
        raise NotImplementedError('Need to be implemented!')

    def getDocument(self, url=None):
        raise NotImplementedError('Need to be implemented!')

    def isDisposed(self):
        return self._disposed

    def getDocumentSubject(self, document):
        return document.DocumentProperties.Subject

    def saveDocumentAs(self, document, format):
        url = None
        name, extension = self._getNamedExtension(document.Title)
        if extension is None:
            extension = self._getDocumentExtension(document)
        filter = self._getDocumentFilter(extension, format)
        if filter is not None:
            temp = getPathSettings(self._ctx).Temp
            url = '%s/%s.%s' % (temp, name, format)
            descriptor = getPropertyValueSet({'FilterName': filter, 'Overwrite': True})
            document.storeToURL(url, descriptor)
            url = getUrl(self._ctx, url)
            if url is not None:
                url = url.Main
        return url

    def parseAttachments(self, attachments, pdf):
        urls = []
        transformer = getUrlTransformer(self._ctx)
        for attachment in attachments:
            url = self._parseAttachment(transformer, attachment, pdf)
            urls.append(url)
        return tuple(urls)

    def validateRecipient(self, email, exist):
        return all((self._isEmailValid(email), not exist))

    def removeSender(self, sender):
        return self.DataSource.removeSender(sender)

    def getSenders(self, *args):
        self.DataSource.getSenders(*args)

# MailModel setter methods
    def dispose(self):
        self._disposed = True

    def setUrl(self, url):
        raise NotImplementedError('Need to be implemented!')

    def setDocumentSubject(self, document, subject):
        document.DocumentProperties.Subject = subject

# MailModel StringRessoure methods
    def getFilePickerTitle(self):
        resource = self._resources.get('PickerTitle')
        title = self._resolver.resolveString(resource)
        return title

    def getDocumentMessage(self, document):
        resource = self._resources.get('Document')
        label = self._resolver.resolveString(resource) % document.Title
        return label

    def getDocumentUserProperty(self, document, name):
        resource = self._resources.get('Property') % name
        state = self._getDocumentUserProperty(document, resource)
        return state

    def getDocumentAttachemnts(self, document, name):
        resource = self._resources.get('Property') % name
        attachments = self._getDocumentAttachments(document, resource)
        return attachments

    def setDocumentUserProperty(self, document, name, value):
        resource = self._resources.get('Property') % name
        self._setDocumentUserProperty(document, resource, value)

    def setDocumentAttachments(self, document, name, value):
        resource = self._resources.get('Property') % name
        self._setDocumentAttachments(document, resource, value)

# MailModel private getter methods
    def _getDocumentUserProperty(self, document, resource, default=True):
        name = self._resolver.resolveString(resource)
        properties = document.DocumentProperties.UserDefinedProperties
        if properties.PropertySetInfo.hasPropertyByName(name):
            value = properties.getPropertyValue(name)
        else:
            value = default
        return value

    def _getDocumentAttachments(self, document, resource):
        attachments = []
        value = self._getDocumentUserProperty(document, resource, None)
        if value is not None:
            try:
                attachments = json.loads(value, object_pairs_hook=OrderedDict)
            except:
                pass
        return tuple(attachments)

    def _isEmailValid(self, email):
        if validators.email(email):
            return True
        return False

    def _parseAttachment(self, transformer, attachment, pdf):
        url = parseUrl(transformer, attachment)
        if pdf:
            self._addPdfMark(url)
        return transformer.getPresentation(url, False)

    def _getNamedExtension(self, name):
        part1, dot, part2 = name.rpartition('.')
        if dot:
            name, extension = part1, part2
        else:
            name, extension = part2, None
        return name, extension

    def _getDocumentExtension(self, document):
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

    def _hasPdfFilter(self, extension):
        filter = self._getDocumentFilter(extension, 'pdf')
        return filter is not None

    def _getDocumentFilter(self, extension, format):
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

# MailModel private setter methods
    def _setDocumentAttachments(self, document, resource, values):
        value = json.dumps(values) if len(values) else None
        self._setDocumentUserProperty(document, resource, value)

    def _setDocumentUserProperty(self, document, resource, value):
        name = self._resolver.resolveString(resource)
        properties = document.DocumentProperties.UserDefinedProperties
        if properties.PropertySetInfo.hasPropertyByName(name):
            if value is None:
                properties.removeProperty(name)
            else:
                properties.setPropertyValue(name, value)
        elif value is not None:
            properties.addProperty(name,
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.MAYBEVOID") +
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.BOUND") +
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.REMOVABLE") +
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.MAYBEDEFAULT"),
            value)

    def _addPdfMark(self, url):
        name, extension = self._getNamedExtension(url.Name)
        if self._hasPdfFilter(extension):
            url.Complete += '#pdf'
