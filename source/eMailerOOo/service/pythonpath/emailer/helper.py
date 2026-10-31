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

from com.sun.star.awt.MessageBoxType import ERRORBOX

from com.sun.star.text.MailMergeType import FILE

from com.sun.star.sdb.CommandType import TABLE

from .unotool import createMessageBox
from .unotool import createService
from .unotool import executeDesktopDispatch
from .unotool import executeFrameDispatch
from .unotool import getDocument
from .unotool import getFileSequence
from .unotool import getInteractionHandler
from .unotool import getLastNamedParts
from .unotool import getNamedValueSet
from .unotool import getMailMerge
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
import traceback


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

def getFilteredRowSet(rowset, filter):
    # XXX: If RowSet.ApplyFilter is not disabled then enabled, RowSet.RowCount is always 1
    rowset.ApplyFilter = False
    rowset.Filter = filter
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
    lenght, sequence = getFileSequence(ctx, url)
    img = base64.b64encode(sequence.value).decode('utf-8')
    return img

def hasExtensionFilter(extension, format):
    return getExtensionFilter(extension, format) is not None

def hasDocumentFilter(document, format):
    return hasExtensionFilter(getDocumentExtension(document), format)

def getMailService(ctx, stype):
    service = 'com.sun.star.mail.MailServiceProvider'
    mailtype = uno.Enum('com.sun.star.mail.MailServiceType', stype)
    return createService(ctx, service).create(mailtype)

def getMailServer(ctx, user,  mailtype):
    server = getMailService(ctx, mailtype.value)
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

def getMailSender(ctx):
    service = 'com.sun.star.mail.MailSender'
    sender = createService(ctx, service)
    return sender

def getMailMessage(ctx, sender, recipient, subject, body):
    service = 'com.sun.star.mail.MailMessage'
    mail = createService(ctx, service, recipient, sender, subject, body)
    return mail

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

def saveDocumentTo(document, path, title, filter, export=None):
    suffix = ''
    descriptor = {'Overwrite': True}
    name, ext = getLastNamedParts(title)
    if filter and hasDocumentFilter(document, filter):
        suffix += '.' + filter
        descriptor['FilterName'] = getDocumentFilter(document, filter)
        if filter == 'pdf' and export:
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

def getDocumentFilter(document, format):
    filter = ''
    extension = getDocumentExtension(document)
    if extension:
        filter = getExtensionFilter(extension, format)
    return filter

def getExtensionFilter(extension, format):
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
    filter = filters.get(format, None)
    return filter

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
    type = infos['type']

    if type == 'datasource':
        master.DataBaseName = datasource
        return True

    elif type == 'table':
        master.DataTableName = table
        return True

    elif type == 'column':
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

def _getExceptionMessage(logger, code, extension, *args):
    title = logger.resolveString(code, extension)
    message = logger.resolveString(code + 1, *args)
    return title, message

def _showWarning(ctx, title, msg):
    msgbox = createMessageBox(ctx, title, msg, ERRORBOX)
    msgbox.execute()
    msgbox.dispose()

def _getInvalidFields(job):
    fields = None
    connection = job.getPropertyValue('ActiveConnection')
    datasource = job.getPropertyValue('DataSourceName')
    document = job.getPropertyValue('Model')
    table = job.getPropertyValue('Command')
    columns = connection.getTables().getByName(table).getColumns().getElementNames()
    masters = document.getTextFieldMasters()
    for name in masters.getElementNames():
        field = masters.getByName(name)
        fields = _getMissingFields(field, document.Title, datasource, table, columns)
        if fields:
            break
    return fields

def _getMissingFields(field, title, datasource, table, columns):
    fields = name = None
    properties = field.getPropertySetInfo()
    if _checkProperty(field, properties, 'DataBaseName', datasource):
        name = 'DataBaseName'
    elif _checkProperty(field, properties, 'DataTableName', table):
        name = 'DataTableName'
    elif _checkProperties(field, properties, 'DataColumnName', columns):
        name = 'DataColumnName'
    if name:
        fields = (title, field.getPropertyValue(name), name, datasource)
    return fields

def _checkProperty(field, properties, name, value):
    return properties.hasPropertyByName(name) and field.getPropertyValue(name) != value

def _checkProperties(field, properties, name, values):
    return properties.hasPropertyByName(name) and field.getPropertyValue(name) not in values

def _getDescriptor(task):
    name, extension = getLastNamedParts(task.Name)
    descriptor = {'OutputURL': task.Folder,
                  'FileNamePrefix': name}
    suffix = ''
    if task.Filter and hasExtensionFilter(extension, task.Filter):
        descriptor['SaveFilter'] = getExtensionFilter(extension, task.Filter)
        suffix += '.' + task.Filter
    elif extension:
        suffix += '.' + extension
    task.Url = task.Folder + g_sep + name + '%s' + suffix
    task.Name = name + suffix
    return getNamedValueSet(descriptor)

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

