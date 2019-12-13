#!
# -*- coding: utf_8 -*-

import uno
import unohelper

from com.sun.star.lang import XServiceInfo
from com.sun.star.task import XJobExecutor, XInteractionHandler
from com.sun.star.awt import XDialogEventHandler, XActionListener, XItemListener
from com.sun.star.sdbc import XRowSetListener
from com.sun.star.uno import XCurrentContext
from com.sun.star.mail import XAuthenticator, MailAttachment
from com.sun.star.datatransfer import XTransferable, DataFlavor
from com.sun.star.beans import PropertyValue
from com.sun.star.util import URL


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = "com.gmail.prrvchr.extensions.smtpMailerOOo.Dialog"


class PyDialog(unohelper.Base,
               XServiceInfo,
               XJobExecutor,
               XDialogEventHandler,
               XActionListener,
               XItemListener,
               XRowSetListener,
               XCurrentContext,
               XAuthenticator,
               XTransferable,
               XInteractionHandler):
    def __init__(self, ctx):
        self.ctx = ctx
        self.tableindex = 0
        self.columnindex = 4
        self.columnfilters = (0, 1, 4)
        self.attachmentsjoin = ", "
        self.configuration = "com.gmail.prrvchr.extensions.smtpMailerOOo/Options"
        self.maxsize = self._getConfiguration(self.configuration).getByName("MaxSizeMo") * 1024 * 1024
        self.document = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx).CurrentComponent
        self.services = ("com.sun.star.text.TextDocument",
                         "com.sun.star.sheet.SpreadsheetDocument",
                         "com.sun.star.drawing.DrawingDocument",
                         "com.sun.star.presentation.PresentationDocument")
        resource = "com.sun.star.resource.StringResourceWithLocation"
        arguments = (self._getResourceLocation(), True, self._getCurrentLocale(), "DialogStrings", "", self)
        self.resource = self.ctx.ServiceManager.createInstanceWithArgumentsAndContext(resource, arguments, self.ctx)
        self.dialog = None
        self.table = None
        self.query = None
        self.address = None
        self.recipient = None
        self.transferable = []
        self.index = -1

    # XJobExecutor
    def trigger(self, args):
        provider = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider", self.ctx)
        self.dialog = provider.createDialogWithHandler("vnd.sun.star.script:smtpMailerOOo.Dialog?location=application", self)
        configuration = self._getConfiguration(self.configuration)
        datasources = self._getDataSources()
        if not self._checkDocumentService():
            self._logMessage("TextMerge", "79.Dialog.TextMerge.Text", ("",))
            self._openDialog(5)
        elif not configuration.getByName("OffLineUse") and not self._checkMailService():
            self._openDialog(5)
        elif not len(datasources):
            self._logMessage("TextMsg", "12.Dialog.TextMsg.Text")
            self._openDialog(1)
        else:
            self.dialog.getControl("AddressBook").addActionListener(self)
            self.dialog.getControl("Address").addItemListener(self)
            self.dialog.getControl("Recipient").addItemListener(self)
            self.dialog.getControl("Attachments").addItemListener(self)
            self.address = self.ctx.ServiceManager.createInstance("com.sun.star.sdb.RowSet")
            self.address.CommandType = uno.getConstantByName("com.sun.star.sdb.CommandType.TABLE")
            self.address.addRowSetListener(self)
            self.recipient = self.ctx.ServiceManager.createInstance("com.sun.star.sdb.RowSet")
            self.recipient.CommandType = uno.getConstantByName("com.sun.star.sdb.CommandType.QUERY")
            self.recipient.addRowSetListener(self)
            control = self.dialog.getControl("DataSource")
            control.addActionListener(self)
            control.Model.StringItemList = datasources
            datasource = self._getDocumentDataSource()
            if datasource in datasources:
                control.selectItem(datasource, True)
            else:
                control.selectItemPos(0, True)
            self._setStep()
            self.dialog.getControl("Subject").Text = self.document.DocumentProperties.Subject
            description = self.dialog.getControl("Description")
            state = self._getDocumentUserProperty("SendAsHtml", configuration.getByName("SendAsHtml"))
            description.Text = self.document.DocumentProperties.Description
            description.Model.Enabled = not state
            self.dialog.getControl("SendAsHtml").Model.State = state
            self.dialog.getControl("SendAsPdf").Model.State = self._getDocumentUserProperty("SendAsPdf", configuration.getByName("SendAsPdf"))
            self.dialog.getControl("Attachments").Model.StringItemList = self._getAttachmentsPath("")
            self.dialog.execute()
            self._saveDialog()
            self.recipient.dispose()
            self.address.dispose()
            self.dialog.getControl("Attachments").dispose()
            self.dialog.getControl("Recipient").dispose()
            self.dialog.getControl("Address").dispose()
            self.dialog.getControl("AddressBook").dispose()
            self.dialog.getControl("DataSource").dispose()
            self.dialog.dispose()
            self.dialog = None

    # XInteractionHandler
    def handle(self, request):
        pass

    # XCurrentContext
    def getValueByName(self, name):
        path = "/org.openoffice.Office.Writer/MailMergeWizard"
        if name == "ServerName":
            if self._getMailServiceType() == uno.Enum("com.sun.star.mail.MailServiceType", "SMTP"):
                return self._getConfiguration(path).getByName("MailServer")
            else:
                return self._getConfiguration(path).getByName("InServerName")
        elif name == "Port":
            if self._getMailServiceType() == uno.Enum("com.sun.star.mail.MailServiceType", "SMTP"):
                return self._getConfiguration(path).getByName("MailPort")
            else:
                return self._getConfiguration(path).getByName("InServerPort")
        elif name ==  "ConnectionType":
            if self._getConfiguration(path).getByName("IsSecureConnection"):
                return "SSL"
            else:
                return "Insecure"

    # XAuthenticator
    def getUserName(self):
        path = "/org.openoffice.Office.Writer/MailMergeWizard"
        if self._getMailServiceType() == uno.Enum("com.sun.star.mail.MailServiceType", "SMTP"):
            return self._getConfiguration(path).getByName("MailUserName")
        else:
            return self._getConfiguration(path).getByName("InServerUserName")
    def getPassword(self):
        path = "/org.openoffice.Office.Writer/MailMergeWizard"
        if self._getMailServiceType() == uno.Enum("com.sun.star.mail.MailServiceType", "SMTP"):
            return self._getConfiguration(path).getByName("MailPassword")
        else:
            return self._getConfiguration(path).getByName("InServerPassword")

    # XTransferable
    def getTransferData(self, flavor):
        transferable = self.transferable.pop(0)
        if transferable == "body":
            if flavor.MimeType == "text/plain;charset=utf-16":
                return self.document.DocumentProperties.Description
            elif flavor.MimeType == "text/html;charset=utf-8":
                return self._getUrlContent(self._saveDocumentAs("html"))
        else:
            return self._getUrlContent(transferable)
    def getTransferDataFlavors(self):
        flavor = DataFlavor()
        transferable = self.transferable[0]
        if transferable == "body":
            if self._getDocumentUserProperty("SendAsHtml"):
                flavor.MimeType = "text/html;charset=utf-8"
                flavor.HumanPresentableName = "HTML-Documents"
            else:
                flavor.MimeType = "text/plain;charset=utf-16"
                flavor.HumanPresentableName = "Unicode text"
        else:
            type = self._getAttachmentType(transferable)
            flavor.MimeType = type
            flavor.HumanPresentableName = type
        return (flavor,)
    def isDataFlavorSupported(self, flavor):
        transferable = self.transferable[0]
        if transferable == "body":
            return (flavor.MimeType == "text/plain;charset=utf-16" or flavor.MimeType == "text/html;charset=utf-8")
        else:
            return flavor.MimeType == self._getAttachmentType(transferable)

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        if self.dialog == dialog:
            if method == "DialogBack":
                self._setStep(event.Source.Model.Tag)
                return True
            elif method == "DialogNext":
                self._setStep(event.Source.Model.Tag)
                return True
            elif method == "AddressBook":
                self._executeShell("thunderbird", "-addressbook")
                return True
            elif method == "RecipientAdd":
                recipients = self._getRecipientFilters()
                filters = self._getAddressFilters(self.dialog.getControl("Address").Model.SelectedItems, recipients)
                self._rowRecipientExecute(recipients + filters)
                return True
            elif method == "RecipientAddAll":
                recipients = self._getRecipientFilters()
                filters = self._getAddressFilters(range(self.address.RowCount), recipients)
                self._rowRecipientExecute(recipients + filters)
                return True
            elif method == "RecipientRemove":
                recipients = self._getRecipientFilters(self.dialog.getControl("Recipient").Model.SelectedItems)
                self._rowRecipientExecute(recipients)
                return True
            elif method == "RecipientRemoveAll":
                self._rowRecipientExecute([])
                return True
            elif method == "SendAsHtml":
                state = (event.Source.Model.State == 1)
                self._setDocumentUserProperty("SendAsHtml", state)
                self.dialog.getControl("Description").Model.Enabled = not state
                return True
            elif method == "ViewHtml":
                url = self._saveDocumentAs("html")
                self._executeShell(url)
                return True
            elif method == "SendAsPdf":
                state = (event.Source.Model.State == 1)
                self._setDocumentUserProperty("SendAsPdf", state)
                return True
            elif method == "ViewPdf":
                url = self._saveDocumentAs("pdf")
                self._executeShell(url)
                return True
            elif method == "AttachmentAdd":
                control = self.dialog.getControl("Attachments")
                for url in self._getSelectedFiles():
                    path = uno.fileUrlToSystemPath(url)
                    if path not in control.Model.StringItemList:
                        control.addItem(path, control.ItemCount)
                self._setAttachments(control.SelectedItemPos)
                return True
            elif method == "AttachmentRemove":
                control = self.dialog.getControl("Attachments")
                for i in reversed(control.Model.SelectedItems):
                    control.removeItems(i, 1)
                self._setAttachments(control.SelectedItemPos)
                return True
            elif method == "AttachmentView":
                url = uno.systemPathToFileUrl(self.dialog.getControl("Attachments").SelectedItem)
                self._executeShell(url)
                return True
        return False
    def getSupportedMethodNames(self):
        return ("DialogBack", "DialogNext", "AddressBook", "RecipientAdd", "RecipientAddAll",
                "RecipientRemove", "RecipientRemoveAll", "SendAsHtml", "ViewHtml", "SendAsPdf",
                "ViewPdf", "AttachmentAdd", "AttachmentRemove", "AttachmentView")

    # XActionListener, XItemListener, XRowSetListener
    def disposing(self, event):
        if event.Source == self.recipient:
            event.Source.removeRowSetListener(self)
            self.recipient = None
        elif event.Source == self.address:
            event.Source.removeRowSetListener(self)
            self.address = None
        elif event.Source == self.dialog.getControl("Attachments"):
            event.Source.removeItemListener(self)
        elif event.Source == self.dialog.getControl("Recipient"):
            event.Source.removeItemListener(self)
        elif event.Source == self.dialog.getControl("Address"):
            event.Source.removeItemListener(self)
        elif event.Source == self.dialog.getControl("AddressBook"):
            event.Source.removeActionListener(self)
        elif event.Source == self.dialog.getControl("DataSource"):
            event.Source.removeActionListener(self)
    def actionPerformed(self, event):
        if event.Source == self.dialog.getControl("DataSource"):
            datasource = event.Source.SelectedItem
            databases = self.ctx.ServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
            if databases.hasByName(datasource):
                database = databases.getByName(datasource)
                if database.IsPasswordRequired:
                    self._logMessage("TextMsg", "15.Dialog.TextMsg.Text", (datasource,))
                else:
                    connection = database.getConnection("","")
                    self._logMessage("TextMsg", "16.Dialog.TextMsg.Text", (datasource,))
                    self.table = connection.Tables.getByIndex(self.tableindex)
                    self.address.DataSourceName = datasource
                    self.address.Filter = self._getQueryFilter()
                    self.address.ApplyFilter = True
                    self.address.Order = self._getQueryOrder()
                    self.query = self._getQuery(database)
                    self._setDocumentDataSource(datasource, self.query.Name)
                    tables = connection.Tables.ElementNames
                    control = self.dialog.getControl("AddressBook")
                    control.Model.StringItemList = tables
                    if self.query.UpdateTableName in tables:
                        control.selectItem(self.query.UpdateTableName, True)
                    elif control.ItemCount:
                        control.selectItemPos(0, True)
                    self.recipient.DataSourceName = datasource
                    self.recipient.Command = self.query.Name
                    self.recipient.Filter = self.query.Filter
                    self.recipient.ApplyFilter = True
                    self.recipient.Order = self.query.Order
                    self.recipient.execute()
                    connection.close()
                    connection.dispose()
                    self._logMessage("TextMsg", "17.Dialog.TextMsg.Text")
        elif event.Source == self.dialog.getControl("AddressBook"):
            self.query.UpdateTableName = event.Source.SelectedItem
            self.address.Command = event.Source.SelectedItem
            self.address.execute()
    def itemStateChanged(self, event):
        if event.Source == self.dialog.getControl("Address"):
            self.dialog.getControl("ButtonAdd").Model.Enabled = (event.Source.SelectedItemPos != -1)
        elif event.Source == self.dialog.getControl("Recipient"):
            position = event.Source.SelectedItemPos
            self.dialog.getControl("ButtonRemove").Model.Enabled = (position != -1)
            if position != -1 and position != self.index:
                self._setDocumentRecord(position)
        elif event.Source == self.dialog.getControl("Attachments"):
            self._setAttachments(event.Source.SelectedItemPos)
    def cursorMoved(self, event):
        pass
    def rowChanged(self, event):
        pass
    def rowSetChanged(self, event):
        if event.Source == self.address:
            self.dialog.getControl("ButtonAdd").Model.Enabled = False
            self.dialog.getControl("Address").Model.StringItemList = self._getRowResult(self.address)
            self.dialog.getControl("ButtonAddAll").Model.Enabled = (self.address.RowCount != 0)
        elif event.Source == self.recipient:
            self.dialog.getControl("ButtonRemove").Model.Enabled = False
            self.dialog.getControl("Recipient").Model.StringItemList = self._getRowResult(self.recipient)
            self.dialog.getControl("ButtonRemoveAll").Model.Enabled = (self.recipient.RowCount != 0)
            if self.dialog.Model.Step == 2:
                self.dialog.getControl("ButtonNext").Model.Enabled = (self.recipient.RowCount != 0)

    # XServiceInfo
    def supportsService(self, service):
        return g_ImplementationHelper.supportsService(g_ImplementationName, service)
    def getImplementationName(self):
        return g_ImplementationName
    def getSupportedServiceNames(self):
        return g_ImplementationHelper.getSupportedServiceNames(g_ImplementationName)

    def _openDialog(self, step):
        self.dialog.Model.Step = step
        self._setDialogError()
        self.dialog.execute()
        self.dialog.dispose()
        self.dialog = None

    def _setDialogError(self):
        self.dialog.Title = self.resource.resolveString("7.Dialog.Title")
        next = self.dialog.getControl("ButtonNext").Model
        next.Label = self.resource.resolveString("23.Dialog.ButtonNext.Label")
        next.Tag = "7"
        next.Enabled = True

    def _checkDocumentService(self):
        for service in self.services:
            if self.document.supportsService(service):
                return True
        return False

    def _getAttachments(self, default=None):
        attachments = self._getDocumentUserProperty("Attachments", default)
        attachments = attachments.split(self.attachmentsjoin) if attachments else []
        return attachments

    def _getAttachmentsPath(self, default=None):
        paths = []
        for url in self._getAttachments(default):
            paths.append(uno.fileUrlToSystemPath(url))
        return paths

    def _getAttachmentsString(self):
        urls = []
        for path in self.dialog.getControl("Attachments").Model.StringItemList:
            urls.append(uno.systemPathToFileUrl(path))
        return self.attachmentsjoin.join(urls)

    def _getDataSources(self, url="sdbc:address:thunderbird"):
        datasources = []
        context = self.ctx.ServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
        for datasource in context.ElementNames:
            if context.getByName(datasource).URL == url:
                id = "13.Dialog.TextMsg.Text"
                datasources.append(datasource)
            else:
                id = "14.Dialog.TextMsg.Text"
            self._logMessage("TextMsg", id, (datasource,))
        return datasources

    def _getColumn(self, index=None):
        index = self.columnindex if index is None else index
        return self.table.Columns.getByIndex(index)

    def _getQuery(self, database, queryname="smtpMailerOOo"):
        queries = database.QueryDefinitions
        if queries.hasByName(queryname):
            query = queries.getByName(queryname)
        else:
            query = self.ctx.ServiceManager.createInstance("com.sun.star.sdb.QueryDefinition")
            query.Command = "SELECT * FROM \"%s\"" % (self.table.Name)
            query.Filter = self._getQueryFilter([])
            query.Order = self._getQueryOrder()
            query.UpdateTableName = self.table.Name
            query.ApplyFilter = True
            queries.insertByName(queryname, query)
            database.DatabaseDocument.store()
        return query

    def _getQueryFilter(self, filters=None):
        if filters is None:
            result = "(%s)" % (self._getFilter())
        elif len(filters):
            result = "(%s)" % (" OR ".join(filters))
        else:
            result = "(%s AND \"%s\" = '')" % (self._getFilter(), self._getColumn().Name)
        return result

    def _getFilter(self):
        filter = []
        for column in self.columnfilters:
            filter.append("\"%s\" IS NOT NULL" % (self._getColumn(column).Name))
        return " AND ".join(filter)

    def _getQueryOrder(self):
        order = "\"%s\"" % (self._getColumn().Name)
        return order

    def _getRowResult(self, rowset, column=None):
        column = self.columnindex + 1 if column is None else column
        result = []
        rowset.beforeFirst()
        while rowset.next():
            result.append(rowset.getString(column))
        return result

    def _getRecipientFilters(self, rows=[]):
        result = []
        self.recipient.beforeFirst()
        while self.recipient.next():
            row = self.recipient.Row -1
            if row not in rows:
                result.append(self._getFilters(self.recipient))
        return result

    def _getAddressFilters(self, rows, filters=[]):
        result = []
        self.address.beforeFirst()
        for row in rows:
            self.address.absolute(row +1)
            filter = self._getFilters(self.address)
            if filter not in filters:
                result.append(filter)
        return result

    def _getFilters(self, rowset):
        result = []
        for column in self.columnfilters:
            value = rowset.getString(column +1).replace("'", "''")
            result.append("\"%s\" = '%s'" % (self._getColumn(column).Name, value))
        return "(%s)" % (" AND ".join(result))

    def _getDocumentDataSource(self):
        if self.document.supportsService("com.sun.star.text.TextDocument"):
            return self.document.createInstance("com.sun.star.document.Settings").CurrentDatabaseDataSource
        return None

    def _setDocumentDataSource(self, datasource, queryname):
        if self._getDocumentDataSource() != datasource:
            if self.document.supportsService("com.sun.star.text.TextDocument"):
                settings = self.document.createInstance("com.sun.star.document.Settings")
                settings.CurrentDatabaseDataSource = datasource
                fields = self.document.TextFields.createEnumeration()
                while fields.hasMoreElements():
                    field = fields.nextElement()
                    master = field.TextFieldMaster
                    if master.supportsService("com.sun.star.text.fieldmaster.Database"):
                        master.DataBaseName = datasource
                        master.DataTableName = queryname

    def _checkMailService(self):
        self._logMessage("TextMerge", "69.Dialog.TextMerge.Text")
        service = self._getMailService()
        try:
            service.connect(self, self)
        except Exception as e:
            self._logMessage("TextMerge", "71.Dialog.TextMerge.Text", (e,))
        else:
            if service.isConnected():
                service.disconnect()
                self._logMessage("TextMerge", "70.Dialog.TextMerge.Text")
                return True
            else:
                self._logMessage("TextMerge", "71.Dialog.TextMerge.Text", ("",))
        return False

    def _getPageCount(self):
        if self.document.supportsService("com.sun.star.text.TextDocument"):
            return self.document.CurrentController.PageCount
        elif self.document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            return self.document.Sheets.Count
        elif self.document.supportsService("com.sun.star.drawing.DrawingDocument"):
            return self.document.DrawPages.Count
        elif self.document.supportsService("com.sun.star.presentation.PresentationDocument"):
            return self.document.DrawPages.Count

    def _checkAttachments(self):
        state = True
        service = self.ctx.ServiceManager.createInstance("com.sun.star.ucb.SimpleFileAccess")
        for url in self._getAttachments():
            state = state and self._checkAttachment(service, url)
        if state:
            self._logMessage("TextMerge", "76.Dialog.TextMerge.Text", (self._getPageCount(), self.recipient.RowCount))
        return state

    def _checkAttachment(self, service, url):
        self._logMessage("TextMerge", "72.Dialog.TextMerge.Text", (uno.fileUrlToSystemPath(url),))
        if service.exists(url):
            size = service.getSize(url)
            maxsize = self.maxsize if self.maxsize != 0 else size
            if size != 0 and size <= maxsize:
                self._logMessage("TextMerge", "73.Dialog.TextMerge.Text")
                return True
            else:
                self._logMessage("TextMerge", "74.Dialog.TextMerge.Text")
        else:
            self._logMessage("TextMerge", "75.Dialog.TextMerge.Text")
        return False

    def _setStep(self, tag="2"):
        self._saveDialog()
        back = self.dialog.getControl("ButtonBack").Model
        next = self.dialog.getControl("ButtonNext").Model
        if tag == "1":
            back.Enabled = False
            next.Tag = "2"
            next.Enabled = True
            self.dialog.Model.Step = 1
            self.dialog.Title = self.resource.resolveString("1.Dialog.Title")
        elif tag == "2":
            back.Tag = "1"
            back.Enabled = True
            next.Tag = "3"
            next.Enabled = (self.recipient.RowCount != 0)
            self.dialog.Model.Step = 2
            self.dialog.Title = self.resource.resolveString("2.Dialog.Title")
        elif tag == "3":
            back.Tag = "2"
            next.Tag = "4"
            self.dialog.Model.Step = 3
            self.dialog.Title = self.resource.resolveString("3.Dialog.Title")
        elif tag == "4":
            back.Tag = "3"
            next.Tag = "5"
            next.Enabled = True
            next.Label = self.resource.resolveString("21.Dialog.ButtonNext.Label")
            self.dialog.Model.Step = 4
            self.dialog.Title = self.resource.resolveString("4.Dialog.Title")
        elif tag == "5":
            back.Tag = "4"
            next.Tag = "6"
            next.Enabled = False
            next.Label = self.resource.resolveString("22.Dialog.ButtonNext.Label")
            self.dialog.getControl("TextMerge").Text = ""
            self.dialog.Model.Step = 5
            self.dialog.Title = self.resource.resolveString("5.Dialog.Title")
            next.Enabled = self._checkAttachments()
        elif tag == "6":
            next.Tag = "7"
            next.Enabled = False
            if self._sendMessages():
                self.dialog.Title = self.resource.resolveString("6.Dialog.Title")
                next.Label = self.resource.resolveString("23.Dialog.ButtonNext.Label")
                next.Enabled = True
        elif tag == "7":
            self.dialog.endExecute()

    def _saveDialog(self):
        step = self.dialog.Model.Step
        if step == 2 and self.query is not None:
            self.recipient.ActiveConnection.Parent.DatabaseDocument.store()
        elif step == 3:
            self.document.DocumentProperties.Subject = self.dialog.getControl("Subject").Text
            self.document.DocumentProperties.Description =  self.dialog.getControl("Description").Text
        elif step == 4:
            self._setDocumentUserProperty("Attachments", self._getAttachmentsString())

    def _rowRecipientExecute(self, filters):
        self.query.ApplyFilter = False
        self.recipient.ApplyFilter = False
        self.query.Filter = self._getQueryFilter(filters)
        self.recipient.Filter = self.query.Filter
        self.recipient.ApplyFilter = True
        self.query.ApplyFilter = True
        self.recipient.execute()

    def _setAttachments(self, position):
        state = (position != -1)
        self.dialog.getControl("ButtonRemoveAttachment").Model.Enabled = state
        self.dialog.getControl("ButtonViewAttachment").Model.Enabled = state

    def _saveDocumentAs(self, type, url=None):
        url = self._getBaseUrl(type) if url is None else url
        args = []
        value = uno.Enum("com.sun.star.beans.PropertyState", "DIRECT_VALUE")
        args.append(PropertyValue("FilterName", -1, self._getDocumentFilter(type), value))
        args.append(PropertyValue("Overwrite", -1, True, value))
        self.document.storeToURL(url, args)
        return url

    def _getSelectedFiles(self, path=None):
        path = self._getPath().Work if path is None else path
        dialog = self.ctx.ServiceManager.createInstance("com.sun.star.ui.dialogs.FilePicker")
        dialog.setDisplayDirectory(path)
        dialog.setMultiSelectionMode(True)
        if dialog.execute():
            return dialog.SelectedFiles
        return []

    def _getBaseUrl(self, extension):
        url = self._getUrl(self._getPath().Temp).Main
        template = self.document.DocumentProperties.TemplateName
        name = template if template else self.document.Title
        url = "%s/%s.%s" % (url, name, extension)
        return self._getUrl(url).Complete

    def _getDocumentFilter(self, type):
        if self.document.supportsService("com.sun.star.text.TextDocument"):
            filter = {"pdf":"writer_pdf_Export", "html":"XHTML Writer File"}
            return filter[type]
        elif self.document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            filter = {"pdf":"calc_pdf_Export", "html":"XHTML Calc File"}
            return filter[type]
        elif self.document.supportsService("com.sun.star.drawing.DrawingDocument"):
            filter = {"pdf":"draw_pdf_Export", "html":"draw_html_Export"}
            return filter[type]
        elif self.document.supportsService("com.sun.star.presentation.PresentationDocument"):
            filter = {"pdf":"impress_pdf_Export", "html":"impress_html_Export"}
            return filter[type]
        return None

    def _getPath(self):
        return self.ctx.ServiceManager.createInstance("com.sun.star.util.PathSettings")

    def _setDocumentRecord(self, index):
        dispatch = None
        frame = self.document.CurrentController.Frame
        if self.document.supportsService("com.sun.star.text.TextDocument"):
            url = self._getUrl(".uno:DataSourceBrowser/InsertContent")
            dispatch = frame.queryDispatch(url, "_self", uno.getConstantByName("com.sun.star.frame.FrameSearchFlag.SELF"))
        elif self.document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            url = self._getUrl(".uno:DataSourceBrowser/InsertColumns")
            dispatch = frame.queryDispatch(url, "_self", uno.getConstantByName("com.sun.star.frame.FrameSearchFlag.SELF"))
        if dispatch is not None:
            dispatch.dispatch(url, self._getDataDescriptor(index + 1))
        self.index = index

    def _getUrl(self, complete):
        url = URL()
        url.Complete = complete
        dummy, url = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.util.URLTransformer", self.ctx).parseStrict(url)
        return url

    def _getDatabase(self, datasource):
        databases = self.ctx.ServiceManager.createInstance("com.sun.star.sdb.DatabaseContext")
        if databases.hasByName(datasource):
            return databases.getByName(datasource)
        return None

    def _logMessage(self, name, id, messages=()):
        control = self.dialog.getControl(name)
        message = self.resource.resolveString(id) % messages
        control.Text = "%s%s" % (control.Text, message)

    def _getDataDescriptor(self, row):
        args = []
        value = uno.Enum("com.sun.star.beans.PropertyState", "DIRECT_VALUE")
        args.append(PropertyValue("DataSourceName", -1, self.recipient.ActiveConnection.Parent.Name, value))
        args.append(PropertyValue("ActiveConnection", -1, self.recipient.ActiveConnection, value))
        args.append(PropertyValue("Command", -1, self.query.Name, value))
        args.append(PropertyValue("CommandType", -1, uno.getConstantByName("com.sun.star.sdb.CommandType.QUERY"), value))
        args.append(PropertyValue("Cursor", -1, self.recipient, value))
        args.append(PropertyValue("Selection", -1, [row], value))
        args.append(PropertyValue("BookmarkSelection", -1, False, value))
        return args

    def _getConfiguration(self, nodepath, update=False):
        value = uno.Enum("com.sun.star.beans.PropertyState", "DIRECT_VALUE")
        config = self.ctx.ServiceManager.createInstance("com.sun.star.configuration.ConfigurationProvider")
        service = "com.sun.star.configuration.ConfigurationUpdateAccess" if update else "com.sun.star.configuration.ConfigurationAccess"
        argument = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        argument.Name = "nodepath"
        argument.Value = nodepath
        return config.createInstanceWithArguments(service, (argument,))

    def _setDocumentUserProperty(self, property, value):
        properties = self.document.DocumentProperties.UserDefinedProperties
        if properties.PropertySetInfo.hasPropertyByName(property):
            properties.setPropertyValue(property, value)
        else:
            properties.addProperty(property,
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.MAYBEVOID") +
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.BOUND") +
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.REMOVABLE") +
            uno.getConstantByName("com.sun.star.beans.PropertyAttribute.MAYBEDEFAULT"),
            value)

    def _getDocumentUserProperty(self, property, default=None):
        properties = self.document.DocumentProperties.UserDefinedProperties
        if properties.PropertySetInfo.hasPropertyByName(property):
            return properties.getPropertyValue(property)
        elif default is not None:
            self._setDocumentUserProperty(property, default)
        return default

    def _executeShell(self, url, option=""):
        shell = self.ctx.ServiceManager.createInstance("com.sun.star.system.SystemShellExecute")
        shell.execute(url, option, 0)

    def _getUrlContent(self, url):
        dummy = None
        service = self.ctx.ServiceManager.createInstance("com.sun.star.ucb.SimpleFileAccess")
        file = service.openFileRead(url)
        length, content = file.readBytes(dummy, service.getSize(url))
        file.closeInput()
        return uno.ByteSequence(content)

    def _sendMessages(self):
        service = self._getMailService()
        try:
            service.connect(self, self)
        except Exception as e:
            self._logMessage("TextMerge", "71.Dialog.TextMerge.Text", (e,))
            self._setDialogError()
        else:
            if service.isConnected():
                state = True
                frm = self._getConfiguration("/org.openoffice.Office.Writer/MailMergeWizard").getByName("MailAddress")
                subject = self.document.DocumentProperties.Subject
                attachments = self._getAttachments()
                self.recipient.beforeFirst()
                while self.recipient.next() and state:
                    self._setDocumentRecord(self.recipient.Row -1)
                    arguments = (self.recipient.getString(self.columnindex +1), frm, subject, self)
                    state = state and self._sendMessage(service, arguments, attachments)
                service.disconnect()
                return state
            else:
                self._logMessage("TextMerge", "71.Dialog.TextMerge.Text", ("",))
                self._setDialogError()
        return False

    def _sendMessage(self, service, arguments, attachments=[]):
        self.transferable = ["body"]
        mail = self.ctx.ServiceManager.createInstanceWithArgumentsAndContext("com.sun.star.mail.MailMessage", arguments, self.ctx)
        if self._getDocumentUserProperty("SendAsPdf"):
            url = self._saveDocumentAs("pdf")
            self.transferable.append(url)
            mail.addAttachment(MailAttachment(self, uno.fileUrlToSystemPath(self._getUrl(url).Name)))
        for url in attachments:
            self.transferable.append(url)
            mail.addAttachment(MailAttachment(self, uno.fileUrlToSystemPath(self._getUrl(url).Name)))
        try:
            service.sendMailMessage(mail)
        except Exception as e:
            recipients = self._getRecipientFilters(range(self.recipient.Row))
            self._rowRecipientExecute(recipients)
            self._logMessage("TextMerge", "78.Dialog.TextMerge.Text", (arguments[0], e))
            self._setDialogError()
        else:
            self._logMessage("TextMerge", "77.Dialog.TextMerge.Text", (arguments[0],))
            return True
        return False

    def _getAttachment(self, url):
        attachment = MailAttachment()
        attachment.Data = self
        attachment.ReadableName = uno.fileUrlToSystemPath(self._getUrl(url).Name)
        return attachment

    def _getAttachmentType(self, url):
        detection = self.ctx.ServiceManager.createInstance("com.sun.star.document.TypeDetection")
        type = detection.queryTypeByURL(url)
        if detection.hasByName(type):
            types = detection.getByName(type)
            for type in types:
                if type.Name == "MediaType":
                    return type.Value
        return "application/octet-stream"

    def _getMailService(self):
        provider = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.mail.MailServiceProvider", self.ctx)
        return provider.create(self._getMailServiceType())

    def _getMailServiceType(self):
        path = "/org.openoffice.Office.Writer/MailMergeWizard"
        if self._getConfiguration(path).getByName("IsSMPTAfterPOP"):
            if self._getConfiguration(path).getByName("InServerIsPOP"):
                return uno.Enum("com.sun.star.mail.MailServiceType", "POP3")
            else:
                return uno.Enum("com.sun.star.mail.MailServiceType", "IMAP")
        else:
            return uno.Enum("com.sun.star.mail.MailServiceType", "SMTP")

    def _getResourceLocation(self):
        identifier = "com.gmail.prrvchr.extensions.smtpMailerOOo"
        provider = self.ctx.getValueByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
        return "%s/smtpMailerOOo" % (provider.getPackageLocation(identifier))

    def _getCurrentLocale(self):
        parts = self._getConfiguration("/org.openoffice.Setup/L10N").getByName("ooLocale").split("-")
        locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        locale.Language = parts[0]
        if len(parts) == 2:
            locale.Country = parts[1]
        else:
            service = self.ctx.ServiceManager.createInstance("com.sun.star.i18n.LocaleData")
            locale.Country = service.getLanguageCountryInfo(locale).Country
        return locale


g_ImplementationHelper.addImplementation(PyDialog,                                                   # UNO object class
                                         g_ImplementationName,                                       # Implementation name
                                        (g_ImplementationName), )                                    # List of implemented services
