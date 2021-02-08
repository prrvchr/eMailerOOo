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

import validators
import traceback


class MailerModel(unohelper.Base):
    def __init__(self, ctx, datasource, path):
        self._ctx = ctx
        self._datasource = datasource
        self._path = path
        self._url = None
        self._stringResource = getStringResource(ctx, g_identifier, g_extension)

    @property
    def DataSource(self):
        return self._datasource
    @property
    def Path(self):
        return self._path
    @Path.setter
    def Path(self, path):
        self._path = path

    def resolveString(self, resource):
        return self._stringResource.resolveString(resource)

    def getUrl(self):
        return self._url

    def setUrl(self, url):
        self._url = url

    def getSenders(self, *args):
        self.DataSource.getSenders(*args)

    def getDocumentSubject(self, document):
        return document.DocumentProperties.Subject

    def getDocumentAttachments(self, document, resource, default=''):
        attachments = ()
        value = self.getDocumentUserProperty(document, resource, default)
        print("MailerModel.getDocumentAttachments() 1 %s" % value)
        if len(value):
            attachments = tuple(value.split('|'))
        print("MailerModel.getDocumentAttachments() 2 %s" % (attachments, ))
        return attachments
        
    def getDocumentUserProperty(self, document, resource, default=True):
        name = self.resolveString(resource)
        properties = document.DocumentProperties.UserDefinedProperties
        if properties.PropertySetInfo.hasPropertyByName(name):
            value = properties.getPropertyValue(name)
        else:
            value = default
        #elif default is not None:
        #    self._setDocumentUserProperty(name, default)
        return value

    def setDocumentSubject(self, document, subject):
        document.DocumentProperties.Subject = subject

    def setDocumentAttachments(self, document, resource, values):
        value = '|'.join(values)
        self.setDocumentUserProperty(document, resource, value)

    def setDocumentUserProperty(self, document, resource, value):
        name = self.resolveString(resource)
        properties = document.DocumentProperties.UserDefinedProperties
        if properties.PropertySetInfo.hasPropertyByName(name):
            properties.setPropertyValue(name, value)
        else:
            properties.addProperty(name,
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.MAYBEVOID") +
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.BOUND") +
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.REMOVABLE") +
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.MAYBEDEFAULT"),
            value)

    def removeSender(self, sender):
        return self.DataSource.removeSender(sender)

    def isEmailValid(self, email):
        if validators.email(email):
            return True
        return False

    def saveDocumentAs(self, document, format):
        url = None
        name, extension = self.getNameAndExtension(document.Title)
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
        attachments = ()
        service = 'com.sun.star.ui.dialogs.FilePicker'
        filepicker = createService(self._ctx, service)
        filepicker.setDisplayDirectory(self._path)
        title = self.resolveString(resource)
        filepicker.setTitle(title)
        filepicker.setMultiSelectionMode(True)
        if filepicker.execute() == OK:
            attachments = filepicker.getSelectedFiles()
            self._path = filepicker.getDisplayDirectory()
        filepicker.dispose()
        return attachments

    def getNameAndExtension(self, filename):
        part1, sep, part2 = filename.rpartition('.')
        if sep:
            name, extension = part1, part2
        else:
            name, extension = part2, part1
        return name, extension

    def hasPdfFilter(self, extension):
        filter = self._getDocumentFilter(extension, 'pdf')
        return filter is not None

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

    def getDocument(self, url=None):
        if url is None:
            url = self._url
        properties = {'Hidden': True, 'MacroExecutionMode': ALWAYS_EXECUTE_NO_WARN}
        descriptor = getPropertyValueSet(properties)
        document = getDesktop(self._ctx).loadComponentFromURL(url, '_blank', 0, descriptor)
        return document
