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

from com.sun.star.mail import MailSpoolerException

from com.sun.star.text.MailMergeType import FILE

from com.sun.star.sdb.CommandType import TABLE

from com.sun.star.uno import Exception as UnoException

from .unotool import checkVersion
from .unotool import createService
from .unotool import executeFrameDispatch
from .unotool import getDocument
from .unotool import getExtensionVersion
from .unotool import getFileSequence
from .unotool import getInteractionHandler
from .unotool import getLastNamedParts
from .unotool import getNamedValueSet
from .unotool import getMailMerge
from .unotool import getMessageBox
from .unotool import getPropertyValueSet
from .unotool import getSimpleFile
from .unotool import getToolKit
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

def getDataBaseContext(ctx, source, name, resolver, code, *args):
    dbcontext = createService(ctx, 'com.sun.star.sdb.DatabaseContext')
    if not dbcontext.hasByName(name):
        msg = resolver(code, name, *args)
        raise MailSpoolerException(msg, source, ())
    location = dbcontext.getDatabaseLocation(name)
    # FIXME: The location can be an embedded database in odt file
    # FIXME: ie: vnd.sun.star.pkg://file://path/document.odt/EmbeddedDatabase
    # FIXME: eMailerOOo cannot work with such a datasource
    if not location.startswith('file://'):
        msg = resolver(code + 1, location, *args)
        raise MailSpoolerException(msg, source, ())
    # We need to check if the registered datasource has an existing odb file
    if not getSimpleFile(ctx).exists(location):
        msg = resolver(code + 2, location, *args)
        raise MailSpoolerException(msg, source, ())
    return dbcontext, location

def getDataSourceConnection(ctx, name, resolver, resource):
    dbcontext, location = getDataBaseContext(ctx, None, name, resolver, resource)
    datasource = dbcontext.getByName(name)
    if datasource.IsPasswordRequired:
        handler = getInteractionHandler(ctx)
        connection = datasource.getIsolatedConnectionWithCompletion(handler)
    else:
        connection = datasource.getIsolatedConnection('', '')
    return connection

def getDataSource(ctx, name, resolver, resource):
    datasource = None
    dbcontext, location = getDataBaseContext(ctx, None, name, resolver, resource)
    if dbcontext.hasByName(name):
        datasource = dbcontext.getByName(name)
    return datasource

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

def parseUriFragment(factory, url, ismerge=True):
    merge = False
    filter = ''
    uri = factory.parse(url)
    if uri.hasFragment():
        fragment = uri.getFragment()
        if ismerge and fragment.startswith('merge'):
            merge = True
        if fragment.endswith('pdf'):
            filter = 'pdf'
        uri.clearFragment()
    return uri.getUriReference(), merge, filter

def mergeDocument(ctx, document, connection, result, datasource, table, row):
    url = None
    if document.supportsService('com.sun.star.text.TextDocument'):
        url = '.uno:DataSourceBrowser/InsertContent'
    elif document.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
        url = '.uno:DataSourceBrowser/InsertColumns'
    if url:
        descriptor = _getDataDescriptor(connection, result, datasource, table, row)
        frame = document.CurrentController.Frame
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

def getTransferable(ctx):
    service = 'com.sun.star.datatransfer.TransferableFactory'
    transferable = createService(ctx, service)
    return transferable

def saveDocumentTo(document, folder, filter, export=None):
    name, extension = getLastNamedParts(document.Title)
    descriptor = {'Overwrite': True}
    suffix = ''
    if filter and hasDocumentFilter(document, filter):
        suffix += '.' + filter
        descriptor['FilterName'] = getDocumentFilter(document, filter)
        if export:
            descriptor['FilterData'] = export.getDescriptor()
    elif extension:
        suffix += '.' + extension
    name += suffix
    temp = folder + '/' + name
    document.storeToURL(temp, getPropertyValueSet(descriptor))
    return temp, name

def getTransferableMimeValues(detection, descriptor, uiname, mimetype, deep=True):
    itype, descriptor = detection.queryTypeByDescriptor(getPropertyValueSet(descriptor), deep)
    if detection.hasByName(itype):
        for t in detection.getByName(itype):
            if t.Name == 'UIName':
                uiname = t.Value
            elif t.Name == 'MediaType':
                mimetype = t.Value
    return uiname, mimetype

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

def getJob(ctx, connection, datasource, table, result, export, selection=None):
    # XXX: To avoid AttributeError: job has no attribute
    # XXX: OutputType, so we only use XPropertySet here...
    job = getMailMerge(ctx)
    job.setPropertyValue('OutputType', FILE)
    job.setPropertyValue('SaveAsSingleFile', False)
    job.setPropertyValue('ActiveConnection', connection)
    job.setPropertyValue('DataSourceName', datasource)
    job.setPropertyValue('Command', table)
    job.setPropertyValue('CommandType', TABLE)
    job.setPropertyValue('ResultSet', result)
    if selection:
        any = uno.Any('[]any', selection)
        uno.invoke(job, 'setPropertyValue', ('Selection', any))
    if export:
        any = uno.Any('[]com.sun.star.beans.PropertyValue', export.getDescriptor())
        uno.invoke(job, 'setPropertyValue', ('SaveFilterData', any))
    return job

def executeMerge(job, task):
    job.setPropertyValue('DocumentURL', task.Url)
    fields = _getInvalidFields(job)
    if fields is None:
        job.getPropertyValue('ResultSet').beforeFirst()
        job.execute(_getDescriptor(task))
    return fields

def executeExport(ctx, task, export):
    document = getDocument(ctx, task.Url)
    task.Url, task.Name = saveDocumentTo(document, task.Folder, task.Filter, export)
    document.close(True)

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
    task.Url = task.Folder + '/' + name + '%s' + suffix
    task.Name = name + suffix
    return getNamedValueSet(descriptor)

def _getDataDescriptor(connection, result, datasource, table, row):
    # FIXME: We need to provide ActiveConnection, DataSourceName, Command and CommandType parameters,
    # FIXME: but apparently only Cursor, BookmarkSelection and Selection parameters are used!!!
    properties = {'ActiveConnection': connection,
                  'DataSourceName': datasource,
                  'Command': table,
                  'CommandType': TABLE,
                  'Cursor': result,
                  'BookmarkSelection': False,
                  'Selection': (row, )}
    descriptor = getPropertyValueSet(properties)
    return descriptor

