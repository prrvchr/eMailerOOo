<!--
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
-->
# [![eMailerOOo logo][1]][2] Documentation

**Ce [document][3] en français.**

**The use of this software subjects you to our [Terms Of Use][4] and [Data Protection Policy][5].**

# version [1.2.7][6]

## Introduction:

**eMailerOOo** is part of a [Suite][7] of [LibreOffice][8] ~~and/or [OpenOffice][9]~~ extensions allowing to offer you innovative services in these office suites.

This extension allows you to send documents in LibreOffice as an email, possibly by mail merge, to your telephone contacts.

It also provides an API allowing you to **[send emails in BASIC][10]** and supporting the most advanced technologies: OAuth2 protocol, Mozilla IspDB, HTTP instead of SMTP/IMAP for Google servers...  
A macro **SendEmail** for sending emails is provided as an example. If you open a document beforehand, you can launch it by:  
**Tools -> Macros -> Run Macro... -> My Macros -> eMailerOOo -> SendEmail -> Main -> Run**

Being free software I encourage you:
- To duplicate its [source code][11].
- To make changes, corrections, improvements.
- To open [issue][12] if needed.

In short, to participate in the development of this extension.  
Because it is together that we can make Free Software smarter.

___

## Requirement:

The eMailerOOo extension uses the OAuth2OOo extension to work.  
It must therefore meet the [requirement of the OAuth2OOo extension][13].

The eMailerOOo extension uses the jdbcDriverOOo extension to work.  
It must therefore meet the [requirement of the jdbcDriverOOo extension][14].

**On Linux and macOS the Python packages** used by the extension, if already installed, may come from the system and therefore **may not be up to date**.  
To ensure that your Python packages are up to date it is recommended to use the **System Info** option in the extension Options accessible by:  
**Tools -> Options -> Internet -> eMailerOOo -> View log -> System Info**  
If outdated packages appear, you can update them with the command:  
`pip install --upgrade <package-name>`

For more information see: [What has been done for version 1.2.0][15].

___

## Installation:

It seems important that the file was not renamed when it was downloaded.  
If necessary, rename it before installing it.

- [![OAuth2OOo logo][17]][18] Install **[OAuth2OOo.oxt][19]** extension [![Version][20]][19]

    You must first install this extension, if it is not already installed.

- [![jdbcDriverOOo logo][21]][22] Install **[jdbcDriverOOo.oxt][23]** extension [![Version][24]][23]

    This extension is necessary to use HsqlDB version 2.7.2 with all its features.

- If you don't have a datasource, you can:

    - [![vCardOOo logo][25]][26] Install **[vCardOOo.oxt][27]** extension [![Version][28]][27]

        This extension is only necessary if you want to use your contacts present on a [Nextcloud][29] platform as a data source for mailing lists and document merging.

    - [![gContactOOo logo][30]][31] Install **[gContactOOo.oxt][32]** extension [![Version][33]][32]

        This extension is only needed if you want to use your personal phone contacts (Android contact) as a data source for mailing lists and document merging.

    - [![mContactOOo logo][34]][35] Install **[mContactOOo.oxt][36]** extension [![Version][37]][36]

        This extension is only needed if you want to use your Microsoft Outlook contacts as a data source for mailing lists and document merging.

    - [![HyperSQLOOo logo][38]][39] Install **[HyperSQLOOo.oxt][40]** extension [![Version][41]][40]

        This extension is only necessary if you want to use a Calc file as a data source for mailing lists and document merging. See: [How to import data from a Calc file][42].

- ![eMailerOOo logo][43] Install **[eMailerOOo.oxt][44]** extension [![Version][45]][44]

Restart LibreOffice after installation.  
**Be careful, restarting LibreOffice may not be enough.**
- **On Windows** to ensure that LibreOffice restarts correctly, use Windows Task Manager to verify that no LibreOffice services are visible after LibreOffice shuts down (and kill it if so).
- **Under Linux or macOS** you can also ensure that LibreOffice restarts correctly, by launching it from a terminal with the command `soffice` and using the key combination `Ctrl + C` if after stopping LibreOffice, the terminal is not active (no command prompt).

___

## Use:

### Introduction:

To be able to use the email merge feature using mailing lists, it is necessary to have a **datasource** with tables having the following columns:
- One or more email addresses columns. These columns are chosen from a list and if this choice is not unique, then the first non-null email address column will be used.
- One or more primary key column to uniquely identify records, it can be a compound primary key. Supported types are VARCHAR and/or INTEGER, or derived. These columns must be declared with the NOT NULL constraint.

In addition, this **datasource** must have at least one **main table**, including all the records that can be used during the email merge.

If you do not have such a **datasource** then I invite you to install one of the following extensions:
- [vCardOOo][26]. This extension will allow you to use your contacts present on a [Nextcloud][29] platform as a data source.
- [gContactOOo][31]. This extension will allow you to use your Android phone (your phone contacts) as a datasource.
- [mContactOOo][35]. This extension will allow you to use your Microsoft Outlook contacts as a datasource.
- [HyperSQLOOo][39]. This extension will allow you to use a Calc file as a datasource. See: [How to import data from a Calc file][42].

For the first 3 extensions the name of the **main table** can be found (and even changed before any connection) in:  
**Tools -> Options -> Internet -> Extension name -> Main table name**

This mode of use is made up of 3 sections:
- [Merge emails with mailing lists][46].
- [Configure connection][47].
- [Outgoing emails][48].

### Merge emails with mailing lists:

#### Requirement:

To be able to post emails to a mailing list, you must:
- Have a **datasource** as described in the previous introduction.
- Open a **new document** in LibreOffice Writer.

This Writer document can include merge fields (insertable by the command: **Insert -> Field -> More fields -> Database -> Mail merge fields**), this is even necessary if you want to be able to customize the content of the email and any attached files.  
These merge fields should only refer to the **main table** of the **datasource**.

If you want to use an **existing Writer document**, you must also ensure that the **datasource** and the **main table** are attached to the document in: **Tools -> Address Book Source**.

If these recommendations are not followed then **merging of documents will not work** and this silently.

#### Starting the mail merge wizard:

In LibreOffice Writer document go to: **Tools -> Add-Ons -> Sending emails -> Merge a document**

![eMailerOOo Merger screenshot 1][49]

#### Data source selection:

The datasource load for the **Email merging** wizard should appear: 

![eMailerOOo Merger screenshot 2][50]

The following screenshots use the [gContactOOo][31] extension as the **datasource**. If you are using your own **datasource**, it is necessary to adapt the settings in relation to it.

In the following screenshot, we can see that the **datasource** gContactOOo is called: `Addresses` and that in the list of tables the table: `PUBLIC.All my contacts` is selected.

![eMailerOOo Merger screenshot 3][51]

If no mailing list exists, you need to create one, by entering its name and validating with: `ENTER` or the `Add` button.

Make sure when creating the mailing list that the **main table** is always selected in the list of tables.  
If this recommendation is not followed then **merging of documents will not work** and this silently.

![eMailerOOo Merger screenshot 4][52]

Now that your new mailing list is available in the list, you need to select it.

And add the following columns:
- Primary key column: `Uri`
- Email address columns: `HomeEmail`, `WorkEmail` and `OtherEmail`

If several columns of email addresses are selected, then the order becomes relevant since the email will be sent to the first available address.  
In addition, on Recipients selection step of the wizard, in the [Available recipients][53] tab, only records with at least one email address column entered will be listed.  
So make sure you have an address book with at least one of the email address field (Home, Work or Other) entered.

![eMailerOOo Merger screenshot 5][54]

This setting is to be made only for new mailing lists.  
You can now proceed to the next step.

#### Recipients selection:

##### Available recipients:

The recipients are selected using 2 buttons `Add all` and `Add` allowing respectively:
- Either add the group of recipients selected from the `Address book` list. This allows during a mailing, that the modifications of the contents of the group are taken into account. A mailing list only accepts one group.
- Either add the selection, which can be multiple using the `CTRL` key. This selection is immutable regardless of the modification of the address book groups.

![eMailerOOo Merger screenshot 6][55]

Example of multiple selection:

![eMailerOOo Merger screenshot 7][56]

##### Selected recipients:

The recipients are deselected using 2 buttons `Remove all` and `Remove` allowing respectively:
- Either remove the group that has been assigned to this mailing list. This is necessary in order to be able to edit the content of this mailing list again.
- Either remove the selection, which can be multiple using the `CTRL` key. 

![eMailerOOo Merger screenshot 8][57]

If you have selected at least 1 recipient, you can proceed to the next step.

#### Sending options selection:

If this is not already done, you must create a new sender using the `Add` button.

![eMailerOOo Merger screenshot 9][58]

The creation of the new sender is described in the [Configure connection][47] section.

The email must have a subject. It can be saved in the Writer document.  
You can insert merge fields in the email subject. A merge field is composed of an opening brace, the name of the referenced column (case sensitive) and a closing brace (ie: `{ColumnName}`).

![eMailerOOo Merger screenshot 10][59]

The email may optionally have attached files. They can be saved in the Writer document.  
The following screenshot shows 1 attached file which will be merged on the data source then converted to PDF format before being attached to the email.

![eMailerOOo Merger screenshot 11][60]

Make sure to always exit the wizard with the `Finish` button to confirm submitting the send jobs.  
To submit mailing jobs, please follow the section [Outgoing emails][48].

### Configure connection:

#### Starting the connection wizard:

In LibreOffice go to: **Tools -> Add-Ons -> Sending emails -> Configure connection**

![eMailerOOo Ispdb screenshot 1][61]

#### Account selection:

![eMailerOOo Ispdb screenshot 2][62]

#### Find the configuration:

![eMailerOOo Ispdb screenshot 3][63]

#### SMTP configuration:

![eMailerOOo Ispdb screenshot 4][64]

#### IMAP configuration:

![eMailerOOo Ispdb screenshot 5][65]

#### Connection test:

![eMailerOOo Ispdb screenshot 6][66]

Always exit the wizard with the `Finish` button to save the connection settings.

### Outgoing emails:

#### Starting the email spooler:

In LibreOffice go to: **Tools -> Add-Ons -> Sending emails -> Outgoing emails**

![eMailerOOo Spooler screenshot 1][67]

#### List of outgoing emails:

Each send job has 3 different states:
- State **0**: the email is ready for sending.
- State **1**: the email was sent successfully.
- State **2**: An error occurred while sending the email. You can view the error message in the [Spooler activity log][68]. 

![eMailerOOo Spooler screenshot 2][69]

The email spooler is stopped by default. **It must be started with the `Start / Stop` button so that the pending emails are sent**.

#### Spooler activity log:

When the email spooler is started, its activity can be viewed in the activity log.

![eMailerOOo Spooler screenshot 3][70]

___

## Sending email with a LibreOffice macro in Basic:

It is possible to send emails using **macros written in Basic**. Sending an email requires a macro of some 50 lines of code and will support most SMTP/IMAP servers.  
Here is the minimum code needed to send an email with attachments.

```
Sub Main

    Rem Ask the user for a sender's email address.
    sSender = InputBox("Please enter the sender's email address")
    Rem User clicked Cancel.
    if sSender = "" then
        exit sub
    endif

    Rem Ask the user for recipient's email address.
    sRecipient = InputBox("Please enter the recipient's email address")
    Rem User clicked Cancel.
    if sRecipient = "" then
        exit sub
    endif

    Rem Ask the user for email's subject.
    sSubject = InputBox("Please enter the email's subject")
    Rem User clicked Cancel.
    if sSubject = "" then
        exit sub
    endif

    Rem Ask the user for email's content.
    sBody = InputBox("Please enter the email's content")
    Rem User clicked Cancel.
    if sBody = "" then
        exit sub
    endif

    Rem Ok now we have everything, we start sending an email.

    Rem We will use 4 UNO services which are:
    Rem - com.sun.star.mail.MailUser: This is the service which will ensure the correct configuration
    Rem   of SMTP and IMAP servers (we can thank Mozilla for the ISPBD database that I use).
    Rem - com.sun.star.mail.MailServiceProvider: This is the service that allows you to use SMTP and
    Rem   IMAP servers. We will use this service with the help of the previous service.
    Rem - com.sun.star.datatransfer.TransferableFactory: This service is a forge for the creation of
    Rem   Transferable which are the basis of the body of the email as well as these attached files.
    Rem - com.sun.star.mail.MailMessage: This is the service that implements the email message.
    Rem Now that everything is clear we can begin.


    Rem First we create the email.

    Rem This is our Transferable forge, it greatly simplifies the LibreOffice mail API...
    oTransferable = createUnoService("com.sun.star.datatransfer.TransferableFactory")

    Rem oBody is the body of the email. It is created here from a String but could also
    Rem have been created from an InputStream, a file Url (file://...) or a sequence of bytes.
    oBody = oTransferable.getByString(sBody)

    Rem oMail is the email message. It is created from the com.sun.star.mail.MailMessage service.
    Rem It can be created with an attachment with the createWithAttachment() method.
    oMail = com.sun.star.mail.MailMessage.create(sRecipient, sSender, sSubject, oBody)

    Rem Ask the user for the URLs of the attached files.
    oDialog = createUnoService("com.sun.star.ui.dialogs.FilePicker")
    oDialog.setMultiSelectionMode(true)
    if oDialog.execute() = com.sun.star.ui.dialogs.ExecutableDialogResults.OK then
        oFiles() = oDialog.getSelectedFiles()
        Rem These two services are simply used to get a suitable file name.
        oUrlTransformer = createUnoService("com.sun.star.util.URLTransformer")
        oUriFactory = createUnoService("com.sun.star.uri.UriReferenceFactory")
        for i = lbound(oFiles()) To ubound(oFiles())
            oUri = getUri(oUrlTransformer, oUriFactory, oFiles(i))
            oAttachment = createUnoStruct("com.sun.star.mail.MailAttachment")
            Rem ReadableName must be entered. This is the name of the attached file
            Rem as it appears in the email. Here we get the file name.
            oAttachment.ReadableName = oUri.getPathSegment(oUri.getPathSegmentCount() - 1)
            Rem The attachment is retrieved from an Url but same as for oBody
            Rem it can be retrieved from a String, an InputStream or a sequence of bytes.
            oAttachment.Data = oTransferable.getByUrl(oUri.getUriReference())
            oMail.addAttachment(oAttachment)
            next i
    endif
    Rem End of creating the email.


    Rem Now we need to send the email.

    Rem First we create a MailUser from the sender address. This is not necessary the
    Rem sender address but it must follow the rfc822 (ie: my surname <myname@myisp.com>).
    Rem The IspDB Wizard will automatically be launched if this user has never been configured.
    oUser = com.sun.star.mail.MailUser.create(sSender)
    Rem User canceled IspDB Wizard.
    if isNull(oUser) then
        exit sub
    endif

    Rem Now that we have the user we can check if they want to use a Reply-To address.
    if oUser.useReplyTo() then
        oMail.ReplyToAddress = oUser.getReplyToAddress()
    endif
    Rem In the same way I can test if the user has an IMAP configuration with oUser.supportIMAP()
    Rem and then create an email thread if necessary. In this case you must:
    Rem - Construct an email message thread (as done previously for oMail).
    Rem - Create and connect to an IMAP server (as we will do for SMTP).
    Rem - Upload this email to the IMAP server with: oServer.uploadMessage(oServer.getSentFolder(), oMail).
    Rem - Once it has been uploaded, retrieve the MessageId with the oMail.MessageId property.
    Rem - Set the oMail.ThreadId property to MessageId for all subsequent emails.
    Rem Great you have successfully grouped the sending of emails into a thread.

    Rem To send the email we need to create an SMTP server. Here's how to do it:
    SMTP = com.sun.star.mail.MailServiceType.SMTP
    oServer = createUnoService("com.sun.star.mail.MailServiceProvider").create(SMTP)
    Rem Now we connect using the SMTP user's configuration.
    oServer.connect(oUser.getConnectionContext(SMTP), oUser.getAuthenticator(SMTP))
    Rem Well, that's it, we are connected, all we have to do is send the email.
    oServer.sendMailMessage(oMail)
    Rem Don't forget to close the connection.
    oServer.disconnect()
    MsgBox "The email has been sent successfully." & chr(13) & "Its MessageId is: " & oMail.MessageId
    Rem Et voilà, you have it...

End Sub


Function getUri(oUrlTransformer As Variant, oUriFactory As Variant, sUrl As String) As Variant
    oUrl = createUnoStruct("com.sun.star.util.URL")
    oUrl.Complete = sUrl
    oUrlTransformer.parseStrict(oUrl)
    oUri = oUriFactory.parse(oUrlTransformer.getPresentation(oUrl, false))
    getUri = oUri
End Function
```

And there you have it, it only took a few lines of code so enjoy...  
However, this is only an example of popularization, and all the necessary error checks are not in place...

___

## Has been tested with:

* LibreOffice 7.3.7.2 - Lubuntu 22.04 - Python version 3.10.12 - OpenJDK-11-JRE (amd64)

* LibreOffice 7.5.4.2(x86) - Windows 10 - Python version 3.8.16 - Adoptium JDK Hotspot 11.0.19 (under Lubuntu 22.04 / VirtualBox 6.1.38)

* LibreOffice 7.4.3.2(x64) - Windows 10(x64) - Python version 3.8.15  - Adoptium JDK Hotspot 11.0.17 (x64) (under Lubuntu 22.04 / VirtualBox 6.1.38)

* LibreOffice 24.2.1.2 - Lubuntu 22.04

* LibreOffice 24.8.0.3 (x86_64) - Windows 10(x64) - Python version 3.9.19 (under Lubuntu 22.04 / VirtualBox 6.1.38)

* **Does not work with OpenOffice** see [bug 128569][71]. Having no solution, I encourage you to install **LibreOffice**.

I encourage you in case of problem :confused:  
to create an [issue][12]  
I will try to solve it :smile:

___

## Historical:

### What has been done for version 0.0.1:

- Writing an [IspDB][72] or SMTP servers connection configuration wizard allowing:
    - Find the connection parameters to an SMTP server from an email address. Besides, I especially thank Mozilla, for [Thunderbird autoconfiguration database][73] or IspDB, which made this challenge possible...
    - Display the activity of the UNO service `com.sun.star.mail.MailServiceProvider` when connecting to the SMTP server and sending an email.

- Writing an email [Spooler][74] allowing:
    - View the email sending jobs with their respective status.
    - Display the activity of the UNO service `com.sun.star.mail.SpoolerService` when sending emails.
    - Start and stop the spooler service.

- Writing an email [Merger][75] allowing:
    - To create mailing lists.
    - To merge and convert the current document to HTML format to make it the email message.
    - To merge and/or convert in PDF format any possible files attached to the email. 

- Writing a document [Mailer][76] allowing:
    - To convert the document to HTML format to make it the email message.
    - To convert in PDF format any possible files attached to the email.

- Writing a [Grid][77] driven by a `com.sun.star.sdb.RowSet` allowing:
    - To be configurable on the columns to be displayed.
    - To be configurable on the sort order to be displayed.
    - Save the display settings.

### What has been done for version 0.0.2:

- Rewrite of [IspDB][72] or Mail servers connection configuration wizard in order to integrate the IMAP connection configuration.
    - Use of [IMAPClient][78] version 2.2.0: an easy-to-use, Pythonic and complete IMAP client library.
    - Extension of [com.sun.star.mail.*][79] IDL files:
        - [XMailMessage2.idl][80] now supports email threading.
        - The new [XImapService.idl][81] interface allows access to part of the IMAPClient library.

- Rewriting of the [Spooler][82] in order to integrate IMAP functionality such as the creation of a thread summarizing the mailing and grouping all the emails sent.

- Submitting the eMailerOOo extension to Google and obtaining permission to use its GMail API to send emails with a Google account.

### What has been done for version 0.0.3:

- Rewrote the [Grid][77] to allow:
    - Sorting on a column with the integration of the UNO service [SortableGridDataModel][83].
    - To generate the filter of records needed by the service [Spooler][74].
    - Sharing the python module with the [Grid][84] module of the [jdbcDriverOOo][22] extension.

- Rewrote the [Merger][75] to allow:
    - Schema name management in table names to be compatible with version 0.0.4 of [jdbcDriverOOo][22]
    - The creation of a mailing list on a group of the address book and allowing to follow the modification of its content.
    - The use of primary key, which can be composite, supporting [DataType][85] `VARCHAR` and `INTEGER` or derived.
    - A preview of the document with merge fields filled in faster thanks to the [Grid][77].

- Rewrote the [Spooler][74] to allow:
    - The use of new filters supporting composite primary keys provided by the [Merger][75].
    - The use of the new [Grid][77] allowing sorting on a column.

- Many other things...

### What has been done for version 1.0.0:

- The **smtpMailerOOo** extension has been renamed to **eMailerOOo**.

### What has been done for version 1.0.1:

- The absence or obsolescence of the **OAuth2OOo** and/or **jdbcDriverOOo** extensions necessary for the proper functioning of **eMailerOOo** now displays an error message. This is to prevent a malfunction such as [issue #3][86] from recurring...

- The underlying HsqlDB database can be opened in Base with: **Tools -> Options -> Internet -> eMailerOOo -> Database**.

- The **Tools -> Add-Ons** menu now displays correctly based on context.

- Many other things...

### What has been done for version 1.0.2:

- If no configuration is found in the connection configuration wizard (IspDB Wizard) then it is possible to configure the connection manually. See [issue #5][87].

### What has been done for version 1.1.0:

- In the connection configuration wizard (IspDB Wizard) it is now possible to deactivate the IMAP configuration.  
    As a result, this no longer sends a thread (IMAP message) when merging a mailing.  
    In this same wizard, it is now possible to enter an email reply-to address.

- In the email merge wizard, it is now possible to insert merge fields in the subject of the email. See [issue #6][88].  
    In the subject of an email, a merge field is composed of an opening brace, the name of the referenced column (case sensitive) and a closing brace (ie: `{ColumnName}`).  
    When entering the email subject, a syntax error in a merge field will be reported and will prevent the mailing from being submitted.

- It is now possible in the Spooler to view emails in eml format.

- A service [com.sun.star.mail.MailUser][89] now allows access to a connection configuration (SMTP and/or IMAP) from an email address following rfc822.  
    Another service [com.sun.star.datatransfer.TransferableFactory][90] allows, as its name suggests, the creation of [Transferable][91] from a String, a binary sequence, an Url (file://...) or a data stream (InputStream).  
    These two new services greatly simplify the LibreOffice mail API and allow sending emails from Basic. See [Issue #4][92].  
    You will find a Basic macro allowing you to send emails in: **Tools -> Macros -> Edit Macros... -> eMailerOOo -> SendEmail**.

### What has been done for version 1.1.1:

- Support for version **1.2.0** of the **OAuth2OOo** extension. Previous versions will not work with **OAuth2OOo** extension 1.2.0 or higher.

### What has been done for version 1.2.0:

- All Python packages necessary for the extension are now recorded in a [requirements.txt][93] file following [PEP 508][94].
- Now if you are not on Windows then the Python packages necessary for the extension can be easily installed with the command:  
  `pip install requirements.txt`
- Modification of the [Requirement][95] section.

### What has been done for version 1.2.1:

- Fixed a regression allowing errors to be displayed in the Spooler.
- Integration of a fix to workaround the [issue #159988][96].

### What has been done for version 1.2.2:

- The creation of the database, during the first connection, uses the UNO API offered by the jdbcDriverOOo extension since version 1.3.2. This makes it possible to record all the information necessary for creating the database in 5 text tables which are in fact [5 csv files][97].
- The extension will ask you to install the OAuth2OOo and jdbcDriverOOo extensions in versions 1.3.4 and 1.3.2 respectively minimum.
- Many fixes.

### What has been done for version 1.2.3:

- Fixed a regression from version 1.2.2 preventing jobs from being submitted to the email spooler.
- Fixed [issue #7][98] not allowing error messages to be displayed in case of incorrect configuration.

### What has been done for version 1.2.4:

- Updated the [Python decorator][99] package to version 5.1.1.
- Updated the [Python ijson][100] package to version 3.3.0.
- Updated the [Python packaging][101] package to version 24.1.
- Updated the [Python setuptools][102] package to version 72.1.0 in order to respond to the [Dependabot security alert][103].
- Updated the [Python validators][104] package to version 0.33.0.
- The extension will ask you to install the OAuth2OOo and jdbcDriverOOo extensions in versions 1.3.6 and 1.4.2 respectively minimum.

### What has been done for version 1.2.5:

- Updated the [Python setuptools][102] package to version 73.0.1.
- The extension will ask you to install the OAuth2OOo and jdbcDriverOOo extensions in versions 1.3.7 and 1.4.5 respectively minimum.
- Changes to extension options that require a restart of LibreOffice will result in a message being displayed.
- Support for LibreOffice version 24.8.x.

### What has been done for version 1.2.6:

- If a reply address was given then it will be used when generating the eml file by the Spooler.
- The extension will ask you to install the OAuth2OOo and jdbcDriverOOo extensions in versions 1.3.8 and 1.4.6 respectively minimum.
- Modification of the extension options accessible via: **Tools -> Options... -> Internet -> eMailerOOo** in order to comply with the new graphic charter.

### What has been done for version 1.2.7:

- The spooler allows opening sent emails either in the local email client (ie: Thunderbird) or online in your browser for accounts using an API for sending email (ie: Google and Microsoft).
- A new tab has been added to the spooler to allow tracking of mail service activity.
- Connections to Microsoft mail servers, which apparently no longer worked, have been migrated to the Graph API.
- For servers that no longer use SMTP and IMAP protocols and offer a replacement API (ie: Google API and Microsoft Graph):
    - All HTTP request parameters needed to send emails are stored in the LibreOffice configuration files.
    - All data needed to process HTTP responses are stored in the LibreOffice configuration files.

    This should allow implementing a third-party API for sending emails just by modifying the [Options.xcu][105] configuration file.
- To work, these new features require the OAuth2OOo extension in version 1.3.9 minimum.
- The command to open an email in Thunderbird can currently only be changed in the LibreOffice configuration (ie: Tools -> Options -> Advanced -> Open Expert Configuration).
- Non-refresh of scrollbars in multi-column lists (ie: grid) has been fixed and will be available from LibreOffice 24.8.4, see [SortableGridDataModel cannot be notified for changes][106].
- Opening emails in your browser does not work with a Microsoft account, the url allowing this has not yet been found and it seems that it would not be possible (ie: popup must be open by the Outlook window)?
- Many fixes.

### What remains to be done for version 1.2.7:

- Add new languages for internationalization...

- Anything welcome...

[1]: </img/emailer.svg#collapse>
[2]: <https://prrvchr.github.io/eMailerOOo/>
[3]: <https://prrvchr.github.io/eMailerOOo/README_fr>
[4]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/TermsOfUse_en>
[5]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/PrivacyPolicy_en>
[6]: <https://prrvchr.github.io/eMailerOOo/#what-has-been-done-for-version-127>
[7]: <https://prrvchr.github.io/>
[8]: <https://www.libreoffice.org/download/download-libreoffice/>
[9]: <https://www.openoffice.org/download/index.html>
[10]: <https://prrvchr.github.io/eMailerOOo/#sending-email-with-a-libreoffice-macro-in-basic>
[11]: <https://github.com/prrvchr/eMailerOOo>
[12]: <https://github.com/prrvchr/eMailerOOo/issues/new>
[13]: <https://prrvchr.github.io/OAuth2OOo/#requirement>
[14]: <https://prrvchr.github.io/jdbcDriverOOo/#requirement>
[15]: <https://prrvchr.github.io/eMailerOOo/#what-has-been-done-for-version-120>
[17]: <https://prrvchr.github.io/OAuth2OOo/img/OAuth2OOo.svg#middle>
[18]: <https://prrvchr.github.io/OAuth2OOo/>
[19]: <https://github.com/prrvchr/OAuth2OOo/releases/latest/download/OAuth2OOo.oxt>
[20]: <https://img.shields.io/github/v/tag/prrvchr/OAuth2OOo?label=latest#right>
[21]: <https://prrvchr.github.io/jdbcDriverOOo/img/jdbcDriverOOo.svg#middle>
[22]: <https://prrvchr.github.io/jdbcDriverOOo/>
[23]: <https://github.com/prrvchr/jdbcDriverOOo/releases/latest/download/jdbcDriverOOo.oxt>
[24]: <https://img.shields.io/github/v/tag/prrvchr/jdbcDriverOOo?label=latest#right>
[25]: <https://prrvchr.github.io/vCardOOo/img/vCardOOo.svg#middle>
[26]: <https://prrvchr.github.io/vCardOOo/>
[27]: <https://github.com/prrvchr/vCardOOo/releases/latest/download/vCardOOo.oxt>
[28]: <https://img.shields.io/github/v/tag/prrvchr/vCardOOo?label=latest#right>
[29]: <https://en.wikipedia.org/wiki/Nextcloud>
[30]: <https://prrvchr.github.io/gContactOOo/img/gContactOOo.svg#middle>
[31]: <https://prrvchr.github.io/gContactOOo/>
[32]: <https://github.com/prrvchr/gContactOOo/releases/latest/download/gContactOOo.oxt>
[33]: <https://img.shields.io/github/v/tag/prrvchr/gContactOOo?label=latest#right>
[34]: <https://prrvchr.github.io/mContactOOo/img/mContactOOo.svg#middle>
[35]: <https://prrvchr.github.io/mContactOOo/>
[36]: <https://github.com/prrvchr/mContactOOo/releases/latest/download/mContactOOo.oxt>
[37]: <https://img.shields.io/github/v/tag/prrvchr/mContactOOo?label=latest#right>
[38]: <https://prrvchr.github.io/HyperSQLOOo/img/HyperSQLOOo.svg#middle>
[39]: <https://prrvchr.github.io/HyperSQLOOo/>
[40]: <https://github.com/prrvchr/HyperSQLOOo/releases/latest/download/HyperSQLOOo.oxt>
[41]: <https://img.shields.io/github/v/tag/prrvchr/HyperSQLOOo?label=latest#right>
[42]: <https://prrvchr.github.io/HyperSQLOOo/#how-to-import-data-from-a-calc-file>
[43]: <https://prrvchr.github.io/eMailerOOo/img/eMailerOOo.svg#middle>
[44]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/eMailerOOo.oxt>
[45]: <https://img.shields.io/github/downloads/prrvchr/eMailerOOo/latest/total?label=v1.2.7#right>
[46]: <https://prrvchr.github.io/eMailerOOo/#merge-emails-with-mailing-lists>
[47]: <https://prrvchr.github.io/eMailerOOo/#configure-connection>
[48]: <https://prrvchr.github.io/eMailerOOo/#outgoing-emails>
[49]: <img/eMailerOOo-Merger1.png>
[50]: <img/eMailerOOo-Merger2.png>
[51]: <img/eMailerOOo-Merger3.png>
[52]: <img/eMailerOOo-Merger4.png>
[53]: <https://prrvchr.github.io/eMailerOOo/#available-recipients>
[54]: <img/eMailerOOo-Merger5.png>
[55]: <img/eMailerOOo-Merger6.png>
[56]: <img/eMailerOOo-Merger7.png>
[57]: <img/eMailerOOo-Merger8.png>
[58]: <img/eMailerOOo-Merger9.png>
[59]: <img/eMailerOOo-Merger10.png>
[60]: <img/eMailerOOo-Merger11.png>
[61]: <img/eMailerOOo-Ispdb1.png>
[62]: <img/eMailerOOo-Ispdb2.png>
[63]: <img/eMailerOOo-Ispdb3.png>
[64]: <img/eMailerOOo-Ispdb4.png>
[65]: <img/eMailerOOo-Ispdb5.png>
[66]: <img/eMailerOOo-Ispdb6.png>
[67]: <img/eMailerOOo-Spooler1.png>
[68]: <https://prrvchr.github.io/eMailerOOo/#spooler-activity-log>
[69]: <img/eMailerOOo-Spooler2.png>
[70]: <img/eMailerOOo-Spooler3.png>
[71]: <https://bz.apache.org/ooo/show_bug.cgi?id=128569>
[72]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/ispdb>
[73]: <https://wiki.mozilla.org/Thunderbird:Autoconfiguration>
[74]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler>
[75]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/merger>
[76]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/mailer>
[77]: <https://github.com/prrvchr/eMailerOOo/tree/master/uno/lib/uno/grid>
[78]: <https://github.com/mjs/imapclient#readme>
[79]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/idl/com/sun/star/mail>
[80]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailMessage2.idl>
[81]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XImapService.idl>
[82]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/spooler.py>
[83]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/grid/SortableGridDataModel.html>
[84]: <https://github.com/prrvchr/jdbcDriverOOo/tree/master/source/jdbcDriverOOo/service/pythonpath/jdbcdriver/grid>
[85]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/sdbc/DataType.html>
[86]: <https://github.com/prrvchr/eMailerOOo/issues/3>
[87]: <https://github.com/prrvchr/eMailerOOo/issues/5>
[88]: <https://github.com/prrvchr/eMailerOOo/issues/6>
[89]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailUser.idl>
[90]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/datatransfer/XTransferableFactory.idl>
[91]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/datatransfer/XTransferable.html>
[92]: <https://github.com/prrvchr/eMailerOOo/issues/4>
[93]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/requirements.txt>
[94]: <https://peps.python.org/pep-0508/>
[95]: <https://prrvchr.github.io/eMailerOOo/#requirement>
[96]: <https://bugs.documentfoundation.org/show_bug.cgi?id=159988>
[97]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/hsqldb>
[98]: <https://github.com/prrvchr/eMailerOOo/issues/7>
[99]: <https://pypi.org/project/decorator/>
[100]: <https://pypi.org/project/ijson/>
[101]: <https://pypi.org/project/packaging/>
[102]: <https://pypi.org/project/setuptools/>
[103]: <https://github.com/prrvchr/eMailerOOo/security/dependabot/1>
[104]: <https://pypi.org/project/validators/>
[105]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/Options.xcu>
[106]: <https://bugs.documentfoundation.org/show_bug.cgi?id=164040>
