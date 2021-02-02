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

from com.sun.star.document.MacroExecMode import ALWAYS_EXECUTE_NO_WARN
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from unolib import getUrl
from unolib import getStringResource
from unolib import getPropertyValueSet
from unolib import getDesktop
from unolib import getPathSettings
from unolib import createService

from smtpserver import g_identifier
from smtpserver import g_extension

from smtpserver import logMessage
from smtpserver import getMessage

from threading import Thread
import validators
import traceback


class MailerModel(unohelper.Base):
    def __init__(self, ctx, datasource):
        self._ctx = ctx
        self._datasource = datasource
        self._document = None
        self._stringResource = getStringResource(ctx, g_identifier, g_extension)

    @property
    def DataSource(self):
        return self._datasource
    @property
    def Document(self):
        return self._document

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

    def getUrl(self):
        return self._document.URL

    def getSenders(self, *args):
        self.DataSource.getSenders(*args)

    def getDocument(self, *args):
        thread = Thread(target=self._getDocument, args=args)
        thread.start()

    def setDocument(self, document):
        self._document = document

    def getDocumentTitle(self, resource):
        title = self.resolveString(resource)
        return title + self._document.URL

    def getDocumentLabel(self, resource):
        label = self.resolveString(resource)
        return label + self._document.Title

    def getDocumentSubject(self):
        return self._document.DocumentProperties.Subject

    def getDocumentDescription(self):
        return self._document.DocumentProperties.Description

    def getDocumentAttachments(self, resource, default=''):
        values = self.getDocumentUserProperty(resource, default)
        attachments = values.split('|')
        return tuple(attachments)
        
    def getDocumentUserProperty(self, resource, default=True):
        name = self.resolveString(resource)
        properties = self._document.DocumentProperties.UserDefinedProperties
        if properties.PropertySetInfo.hasPropertyByName(name):
            value = properties.getPropertyValue(name)
        else:
            value = default
        #elif default is not None:
        #    self._setDocumentUserProperty(name, default)
        return value

    def removeSender(self, sender):
        return self.DataSource.removeSender(sender)

    def isEmailValid(self, email):
        if validators.email(email):
            return True
        return False

    def _getDocument(self, document, url, callback):
        if document is None:
            document = self.loadDocument(url)
        callback(document)

    def loadDocument(self, url):
        properties = {'Hidden': True, 'MacroExecutionMode': ALWAYS_EXECUTE_NO_WARN}
        descriptor = getPropertyValueSet(properties)
        document = getDesktop(self._ctx).loadComponentFromURL(url, '_blank', 0, descriptor)
        return document

    def saveDocumentAs(self, document, format):
        url = None
        name, extension = self._getNameAndExtension(document.Title)
        filter = self._getDocumentFilter(extension, format)
        if filter is not None:
            temp = getPathSettings(self._ctx).Temp
            url = '%s/%s.%s' % (temp, name, format)
            descriptor = getPropertyValueSet({'FilterName': filter, 'Overwrite': True})
            document.storeToURL(url, descriptor)
            url = getUrl(self._ctx, url)
            if url is not None:
                url = url.Main
            print("MailerModel.saveDocumentAs() %s" % url)
        return url

    def getAttachments(self, resource):
        documents = ()
        directory = getPathSettings(self._ctx).Work
        service = 'com.sun.star.ui.dialogs.FilePicker'
        filepicker = createService(self._ctx, service)
        filepicker.setDisplayDirectory(directory)
        title = self.resolveString(resource)
        filepicker.setTitle(title)
        filepicker.setMultiSelectionMode(True)
        if filepicker.execute() == OK:
            documents = filepicker.getFiles()
        filepicker.dispose()
        return documents

    def _getNameAndExtension(self, filename):
        name, sep, extension = filename.rpartition('.')
        return name, extension

    def _getDocumentFilter(self, extension, format):
        if extension == 'odt':
            filters = {'pdf': 'writer_pdf_Export', 'html': 'XHTML Writer File'}
        elif extension == 'ods':
            filters = {'pdf': 'calc_pdf_Export', 'html': 'XHTML Calc File'}
        elif extension == 'odp':
            filters = {'pdf': 'impress_pdf_Export', 'html': 'impress_html_Export'}
        elif extension == 'odg':
            filters = {'pdf': 'draw_pdf_Export', 'html': 'draw_html_Export'}
        else:
            filters = {}
        filter = filters.get(format, None)
        return filter
