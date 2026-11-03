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

from com.sun.star.sdb.CommandType import TABLE

from .unotool import createService
from .unotool import executeDesktopDispatch
from .unotool import executeFrameDispatch
from .unotool import getDocument
from .unotool import getFileSequence
from .unotool import getInteractionHandler
from .unotool import getLastNamedParts
from .unotool import getPropertyValueSet
from .unotool import getSimpleFile

from .oauth20 import g_identifier as g_identifier1
from .oauth20 import g_extension as g_extension1
from .oauth20 import g_version as g_version1

from .jdbcdriver import g_identifier as g_identifier2
from .jdbcdriver import g_extension as g_extension2
from .jdbcdriver import g_version as g_version2

from .configuration import g_sep

import base64


def getExtensions():
    return {g_identifier1: (g_extension1, g_version1),
            g_identifier2: (g_extension2, g_version2)}


def getDataBaseContext(ctx, name):
    dbcontext = createService(ctx, 'com.sun.star.sdb.DatabaseContext')
    if not dbcontext.hasByName(name):
        return 1, dbcontext, ''
    location = dbcontext.getDatabaseLocation(name)
    # FIXME: The location can be an embedded database in odt file
    # FIXME: ie: vnd.sun.star.pkg://file://path/document.odt/EmbeddedDatabase
    # FIXME: eMailerOOo cannot work with such a datasource
    if not location.startswith('file://'):
        return 2, dbcontext, location
    # We need to check if the registered datasource has an existing odb file
    if not getSimpleFile(ctx).exists(location):
        return 3, dbcontext, location
    return 0, dbcontext, location


def getDataSource(ctx, name):
    datasource = None
    code, dbcontext, location = getDataBaseContext(ctx, name)
    if code == 0:
        datasource = dbcontext.getByName(name)
    return code, datasource, location


def getConnection(ctx, datasource, parent=None, context=''):
    if datasource.IsPasswordRequired:
        handler = getInteractionHandler(ctx)
        if parent:
            handler.initialize(getPropertyValueSet({'Parent': parent, 'Context': context}))
        connection = datasource.getIsolatedConnectionWithCompletion(handler)
    else:
        connection = datasource.getIsolatedConnection('', '')
    return connection


def getRowSet(ctx, connection, datasource, table, cmdtype=TABLE):
    service = 'com.sun.star.sdb.RowSet'
    rowset = createService(ctx, service)
    rowset.ActiveConnection = connection
    rowset.DataSourceName = datasource
    rowset.CommandType = cmdtype
    rowset.Command = table
    return rowset


def getFilteredRowSet(rowset, fltr):
    # XXX: If RowSet.ApplyFilter is not disabled then enabled, RowSet.RowCount is always 1
    rowset.ApplyFilter = False
    rowset.Filter = fltr
    rowset.ApplyFilter = True
    rowset.execute()
    return rowset


def parseUri(factory, url):
    uri = factory.parse(url)
    if uri.hasFragment():
        uri.clearFragment()
    return uri.getUriReference()


def parseUriFragment(factory, url, /, merged=True, export=''):
    merge = False
    uri = factory.parse(url)
    if uri.hasFragment():
        fragment = uri.getFragment()
        if fragment.startswith('merge') or export:
            merge = merged
        if fragment.endswith('pdf'):
            export = 'pdf'
        uri.clearFragment()
    return uri.getUriReference(), merge, export


def mergeDocument(ctx, document, connection, datasource, table, resultset, row):
    descriptor = getDataDescriptor(connection, datasource, table, resultset, row)
    executeMerge(ctx, document, getPropertyValueSet(descriptor))


def executeMerge(ctx, document, descriptor):
    frame = document.getCurrentController().getFrame()
    if document.supportsService('com.sun.star.text.TextDocument'):
        url = '.uno:DataSourceBrowser/InsertContent'
    elif document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
        url = '.uno:DataSourceBrowser/InsertColumns'
    else:
        url = '.uno:DataSourceBrowser/DataToFields'
    executeFrameDispatch(ctx, frame, url, None, *descriptor)


def getMessageImage(ctx, url):
    _, sequence = getFileSequence(ctx, url)
    img = base64.b64encode(sequence.value).decode('utf-8')
    return img


def hasExtensionFilter(extension, options):
    return getExtensionFilter(extension, options) is not None


def hasDocumentFilter(document, options):
    return hasExtensionFilter(getDocumentExtension(document), options)


def getMailService(ctx, mailtype):
    service = 'com.sun.star.mail.MailServiceProvider'
    return createService(ctx, service).create(mailtype)


def getMailServer(ctx, user,  mailtype):
    server = getMailService(ctx, mailtype)
    context = user.getConnectionContext(mailtype)
    authenticator = user.getAuthenticator(mailtype)
    server.connect(context, authenticator)
    return server


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


def showMailer(ctx, listener=None, /, **kwargs):
    url = 'vnd.sun.star.job:service=io.github.prrvchr.eMailerOOo.Mailer'
    executeDesktopDispatch(ctx, url, listener, **kwargs)


def showMerger(ctx, listener=None, /, **kwargs):
    url = 'vnd.sun.star.job:service=io.github.prrvchr.eMailerOOo.Merger'
    executeDesktopDispatch(ctx, url, listener, **kwargs)


def showIspDB(ctx, listener=None, /, **kwargs):
    url = 'vnd.sun.star.job:service=io.github.prrvchr.eMailerOOo.IspDB'
    executeDesktopDispatch(ctx, url, listener, **kwargs)


def showController(ctx, listener=None, /, **kwargs):
    url = 'vnd.sun.star.job:service=io.github.prrvchr.eMailerOOo.Controller'
    executeDesktopDispatch(ctx, url, listener, **kwargs)


def showSpooler(ctx, listener=None, /, **kwargs):
    url = 'vnd.sun.star.job:service=io.github.prrvchr.eMailerOOo.Spooler'
    executeDesktopDispatch(ctx, url, listener, **kwargs)


def saveDocumentTo(document, path, title, fltr, export=None):
    suffix = ''
    descriptor = {'Overwrite': True}
    name, ext = getLastNamedParts(title)
    if fltr and hasDocumentFilter(document, fltr):
        suffix += '.' + fltr
        descriptor['FilterName'] = getDocumentFilter(document, fltr)
        if fltr == 'pdf' and export:
            descriptor['FilterData'] = export.getDescriptor()
    elif ext:
        suffix += '.' + ext
    name += suffix
    url = path + g_sep + name
    document.storeToURL(url, getPropertyValueSet(descriptor))
    return url


def getDocumentExtension(document):
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


def getDocumentFilter(document, options):
    fltr = ''
    extension = getDocumentExtension(document)
    if extension:
        fltr = getExtensionFilter(extension, options)
    return fltr


def getExtensionFilter(extension, options):
    if extension:
        extension = extension.lower()
    if extension in ('odt', 'ott', 'odm', 'doc', 'dot'):
        filters = {'pdf': 'writer_pdf_Export', 'html': 'XHTML Writer File'}
    elif extension in ('ods', 'ots', 'xls', 'xlt'):
        filters = {'pdf': 'calc_pdf_Export', 'html': 'XHTML Calc File'}
    elif extension in ('odg', 'otg'):
        filters = {'pdf': 'draw_pdf_Export', 'html': 'draw_html_Export'}
    elif extension in ('odp', 'otp', 'ppt', 'pot'):
        filters = {'pdf': 'impress_pdf_Export', 'html': 'impress_html_Export'}
    else:
        filters = {}
    fltr = filters.get(options, None)
    return fltr


def getInvalidFields(datasource, table, columns, document):
    fields = []
    masters = document.getTextFieldMasters()
    for name in masters.getElementNames():
        if name.startswith('com.sun.star.text.fieldmaster.Database.'):
            master = masters.getByName(name)
            erreur = _analyzeFieldAnomalies(master, datasource, table, columns)
            if erreur:
                fields.append({'master': master,
                               'type': erreur})
    return fields


def _analyzeFieldAnomalies(field, datasource, table, columns):
    if field.DataBaseName != datasource:
        return 'datasource'

    if field.DataTableName != table:
        return 'table'

    if field.DataColumnName not in columns:
        return 'column'
    return None


def tryCorrectFields(infos, datasource, table, columns):
    master = infos['master']
    mtype = infos['type']

    if mtype == 'datasource':
        master.DataBaseName = datasource
        return True

    elif mtype == 'table':
        master.DataTableName = table
        return True

    elif mtype == 'column':
        column = master.DataColumnName
        matches = [c for c in columns if c.lower() == column.lower()]

        if matches:
            master.DataColumnName = matches[0]
            return True
        return False
    return False


def executeExport(ctx, task, export):
    document = getDocument(ctx, task.Url)
    task.Url, task.Name = saveDocumentTo(document, task.Folder, task.Filter, export)
    document.close(True)


def getDataDescriptor(connection, datasource, table, resultset=None, row=1):
    # FIXME: We need to provide ActiveConnection, DataSourceName, Command and CommandType parameters,
    # FIXME: but apparently only Cursor, BookmarkSelection and Selection parameters are used!!!
    return {'ActiveConnection': connection,
            'DataSourceName': datasource,
            'Command': table,
            'CommandType': TABLE,
            'Cursor': resultset,
            'BookmarkSelection': False,
            'Selection': (row, )}
