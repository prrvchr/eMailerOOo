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
import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from ..unotool import createService
from ..unotool import getConfiguration
from ..unotool import getStringResource
from ..unotool import getUrlTransformer
from ..unotool import parseUrl

from ..helper import getDocumentFilter
from ..helper import getNamedExtension
from ..helper import saveDocumentAs

from ..configuration import g_extension
from ..configuration import g_identifier

from collections import OrderedDict
from email.utils import parseaddr
import validators
import json
import traceback


class MailModel(unohelper.Base):
    def __init__(self, ctx, datasource, close=True):
        self._ctx = ctx
        self._datasource = datasource
        self._close = close
        self._config = getConfiguration(ctx, g_identifier, True)
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'MessageBox')
        self._disposed = False

    @property
    def _dataSource(self):
        return self._datasource
    @property
    def Path(self):
        return self._path
    @Path.setter
    def Path(self, path):
        self._path = path
    @property
    def IdentifierQuoteString(self):
        return self._datasource.IdentifierQuoteString

# MailModel getter methods
    def isDisposed(self):
        return self._disposed

    def getUrl(self):
        raise NotImplementedError('Need to be implemented!')

    def getDocument(self, url=None):
        raise NotImplementedError('Need to be implemented!')

    def mergeDocument(self, document, row):
        raise NotImplementedError('Need to be implemented!')

    def isSubjectValid(self, subject):
        raise NotImplementedError('Need to be implemented!')

    def getDocumentSubject(self, document):
        return document.DocumentProperties.Subject

    def saveDocumentAs(self, document, format):
        return saveDocumentAs(self._ctx, document, format)

    def parseAttachments(self, attachments, merge, pdf):
        urls = []
        transformer = getUrlTransformer(self._ctx)
        for attachment in attachments:
            url = self._parseAttachment(transformer, attachment, merge, pdf)
            urls.append(url)
        return tuple(urls)

    def isSenderValid(self, sender):
        name, address = parseaddr(sender)
        return self.isEmailValid(address)

    def isRecipientsValid(self, recipients):
        return len(recipients) > 0

    def isEmailValid(self, email):
        if validators.email(email):
            return True
        return False

    def removeSender(self, sender):
        senders = self._config.getByName('Senders')
        if senders.hasByName(sender):
            senders.removeByName(sender)
            if self._config.hasPendingChanges():
                self._config.commitChanges()
                return True
        return False

    def getSenders(self):
        return self._config.getByName('Senders').ElementNames

# MailModel setter methods
    def dispose(self):
        self._disposed = True
        if self._close:
            self._datasource.dispose()

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

    def _parseAttachment(self, transformer, attachment, merge, pdf):
        url = parseUrl(transformer, attachment)
        if merge or pdf:
            url = self._getUrlParameter(transformer, url, merge, pdf)
        location = transformer.getPresentation(url, False)
        return location

    def _hasPdfFilter(self, extension):
        filter = getDocumentFilter(extension, 'pdf')
        return filter is not None

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

    def _getUrlParameter(self, transformer, url, merge, pdf):
        marks = []
        if merge:
            marks.append('merge')
        name, extension = getNamedExtension(url.Name)
        if pdf and self._hasPdfFilter(extension):
            marks.append('pdf')
        url.Mark = '&'.join(marks)
        success, url = transformer.assemble(url)
        return url
