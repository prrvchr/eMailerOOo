---
layout: default
title: eMailerOOo documentation (English)
permalink: /
redirect_from:
  - /README
  - /README.html
---
<!--
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
-->
# [![eMailerOOo logo][1]][2] Documentation

**Ce [document][3] en français.**

**The use of this software subjects you to our [Terms Of Use][4] and [Data Protection Policy][5].**

# version [1.5.2][6]

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
- To [participate in the costs][13] of [CASA certification][14].

In short, to participate in the development of this extension.  
Because it is together that we can make Free Software smarter.

___

## CASA certification:

To ensure interoperability with **Google**, the **eMailerOOo** extension uses the **OAuth2OOo** extension which requires [CASA certification][14].  
Until now, this certification was free and carried out by a Google partner.  
The **OAuth2OOo** application obtained its [CASA certification][15] on 11/28/2023.

**Now this certification has become paid and costs $600.**

I never anticipated such costs and I am counting on your contribution to finance this certification.

Thank you for your help. [![Sponsor][16]][13]

___

## Requirement:

The eMailerOOo extension uses the OAuth2OOo extension to work.  
It must therefore meet the [requirement of the OAuth2OOo extension][17].

The eMailerOOo extension uses the jdbcDriverOOo extension to work.  
It must therefore meet the [requirement of the jdbcDriverOOo extension][18].  
Additionally, eMailerOOo requires the jdbcDriverOOo extension to be configured to provide `com.sun.star.sdb` as the API level, which is the default configuration.

___

## Installation:

It seems important that the file was not renamed when it was downloaded.  
If necessary, rename it before installing it.

- [![OAuth2OOo logo][19]][20] Install **[OAuth2OOo.oxt][21]** extension [![Version][22]][21]

    You must first install this extension, if it is not already installed.

- [![jdbcDriverOOo logo][23]][24] Install **[jdbcDriverOOo.oxt][25]** extension [![Version][26]][25]

    This extension is necessary to use HsqlDB version 2.7.2 with all its features.

- If you don't have a datasource, you can:

    - [![vCardOOo logo][27]][28] Install **[vCardOOo.oxt][29]** extension [![Version][30]][29]

        This extension is only necessary if you want to use your contacts present on a [Nextcloud][31] platform as a data source for mailing lists and document merging.

    - [![gContactOOo logo][32]][33] Install **[gContactOOo.oxt][34]** extension [![Version][35]][34]

        This extension is only needed if you want to use your personal phone contacts (Android contact) as a data source for mailing lists and document merging.

    - [![mContactOOo logo][36]][37] Install **[mContactOOo.oxt][38]** extension [![Version][39]][38]

        This extension is only needed if you want to use your Microsoft Outlook contacts as a data source for mailing lists and document merging.

    - [![HyperSQLOOo logo][40]][41] Install **[HyperSQLOOo.oxt][42]** extension [![Version][43]][42]

        This extension is only necessary if you want to use a Calc file as a data source for mailing lists and document merging. See: [How to import data from a Calc file][44].

- ![eMailerOOo logo][45] Install **[eMailerOOo.oxt][46]** extension [![Version][47]][46]

Restart LibreOffice after installation.  
**Be careful, restarting LibreOffice may not be enough.**
- **On Windows** to ensure that LibreOffice restarts correctly, use Windows Task Manager to verify that no LibreOffice services are visible after LibreOffice shuts down (and kill it if so).
- **Under Linux or macOS** you can also ensure that LibreOffice restarts correctly, by launching it from a terminal with the command `soffice` and using the key combination `Ctrl + C` if after stopping LibreOffice, the terminal is not active (no command prompt).

After this restart, you will be asked **to install Python packages containing binary files**. Please refer to the [Installing Python packages][48] section for more information.

___

## Use:

### Introduction:

To be able to use the email merge feature using mailing lists, it is necessary to have a **datasource** with tables having the following columns:
- One or more email addresses columns. These columns are chosen from a list and if this choice is not unique, then the first non-null email address column will be used.
- One or more primary key column to uniquely identify records, it can be a compound primary key. Supported types are VARCHAR and/or INTEGER, or derived. These columns must be declared with the NOT NULL constraint.

In addition, this **datasource** must have at least one **main table**, including all the records that can be used during the email merge.

If you do not have such a **datasource** then I invite you to install one of the following extensions:
- [vCardOOo][28]. This extension will allow you to use your contacts present on a [Nextcloud][31] platform as a data source.
- [gContactOOo][33]. This extension will allow you to use your Android phone (your phone contacts) as a datasource.
- [mContactOOo][37]. This extension will allow you to use your Microsoft Outlook contacts as a datasource.
- [HyperSQLOOo][41]. This extension will allow you to use a Calc file as a datasource. See: [How to import data from a Calc file][44].

For the first 3 extensions the name of the **main table** can be found (and even changed before any connection) in:  
**Tools -> Options -> Internet -> Extension name -> Main table name**

This mode of use is made up of 3 sections:
- [Merge emails with mailing lists][49].
- [Configure connection][50].
- [Outgoing emails][51].

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

![eMailerOOo Merger screenshot 1][52]

#### Data source selection:

The datasource load for the **Email merging** wizard should appear: 

![eMailerOOo Merger screenshot 2][53]

The following screenshots use the [gContactOOo][33] extension as the **datasource**. If you are using your own **datasource**, it is necessary to adapt the settings in relation to it.

In the following screenshot, we can see that the **datasource** gContactOOo is called: `Addresses` and that in the list of tables the table: `PUBLIC.All my contacts` is selected.

![eMailerOOo Merger screenshot 3][54]

If no mailing list exists, you need to create one, by entering its name and validating with: `ENTER` or the `Add` button.

Make sure when creating the mailing list that the **main table** is always selected in the list of tables.  
If this recommendation is not followed then **merging of documents will not work** and this silently.

![eMailerOOo Merger screenshot 4][55]

Now that your new mailing list is available in the list, you need to select it.

And add the following columns:
- Primary key column: `Uri`
- Email address columns: `HomeEmail`, `WorkEmail` and `OtherEmail`

If several columns of email addresses are selected, then the order becomes relevant since the email will be sent to the first available address.  
In addition, on Recipients selection step of the wizard, in the [Available recipients][56] tab, only records with at least one email address column entered will be listed.  
So make sure you have an address book with at least one of the email address field (Home, Work or Other) entered.

![eMailerOOo Merger screenshot 5][57]

This setting is to be made only for new mailing lists.  
You can now proceed to the next step.

#### Recipients selection:

##### Available recipients:

The recipients are selected using 2 buttons `Add all` and `Add` allowing respectively:
- Either add the group of recipients selected from the `Address book` list. This allows during a mailing, that the modifications of the contents of the group are taken into account. A mailing list only accepts one group.
- Either add the selection, which can be multiple using the `CTRL` key. This selection is immutable regardless of the modification of the address book groups.

![eMailerOOo Merger screenshot 6][58]

Example of multiple selection:

![eMailerOOo Merger screenshot 7][59]

##### Selected recipients:

The recipients are deselected using 2 buttons `Remove all` and `Remove` allowing respectively:
- Either remove the group that has been assigned to this mailing list. This is necessary in order to be able to edit the content of this mailing list again.
- Either remove the selection, which can be multiple using the `CTRL` key. 

![eMailerOOo Merger screenshot 8][60]

If you have selected at least 1 recipient, you can proceed to the next step.

#### Sending options selection:

If this is not already done, you must create a new sender using the `Add` button.

![eMailerOOo Merger screenshot 9][61]

The creation of the new sender is described in the [Configure connection][50] section.

The email must have a subject. It can be saved in the Writer document.  
You can insert merge fields in the email subject. A merge field is composed of an opening brace, the name of the referenced column (case sensitive) and a closing brace (ie: `{ColumnName}`).

![eMailerOOo Merger screenshot 10][62]

The email may optionally have attached files. They can be saved in the Writer document.  
The following screenshot shows 1 attached file which will be merged on the data source then converted to PDF format before being attached to the email.

![eMailerOOo Merger screenshot 11][63]

Make sure to always exit the wizard with the `Finish` button to confirm submitting the send jobs.  
To submit mailing jobs, please follow the section [Outgoing emails][51].

### Configure connection:

#### Starting the connection wizard:

In LibreOffice go to: **Tools -> Add-Ons -> Sending emails -> Configure connection**

![eMailerOOo Ispdb screenshot 1][64]

#### Account selection:

![eMailerOOo Ispdb screenshot 2][65]

#### Find the configuration:

![eMailerOOo Ispdb screenshot 3][66]

#### SMTP configuration:

![eMailerOOo Ispdb screenshot 4][67]

#### IMAP configuration:

![eMailerOOo Ispdb screenshot 5][68]

#### Connection test:

![eMailerOOo Ispdb screenshot 6][69]

Always exit the wizard with the `Finish` button to save the connection settings.

### Outgoing emails:

#### Starting the email spooler:

In LibreOffice go to: **Tools -> Add-Ons -> Sending emails -> Outgoing emails**

![eMailerOOo Spooler screenshot 1][70]

#### List of outgoing emails:

Each send job has 3 different states:
- State **0**: the email is ready for sending.
- State **1**: the email was sent successfully.
- State **2**: An error occurred while sending the email. You can view the error message in the [Spooler activity log][71]. 

![eMailerOOo Spooler screenshot 2][72]

The email spooler is stopped by default. **It must be started with the `Start / Stop` button so that the pending emails are sent**.

#### Spooler activity log:

When the email spooler is started, its activity can be viewed in the activity log.

![eMailerOOo Spooler screenshot 3][73]

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

## How to customize LibreOffice menus:

To allow you to place access to the various features of eMailerOOo wherever you want, it is now possible to create custom menus for commands:
- `ShowIspdb` to **Configure connection** with the LibreOffice Scope.
- `ShowMailer` to **Send a document** with the LibreOffice Scope.
- `ShowMerger` to **Merge a document** with the Writer and Calc Scope.
- `ShowSpooler` to **Outgoing emails** with the LibreOffice Scope.
- `StartSpooler` to **Start spooler** with the LibreOffice Scope.
- `StopSpooler` to **Stop spooler** with the LibreOffice Scope.

In the **Menu** tab of the **Tools -> Customize** window, select **Macros** in **Category** to access the macros under: **My Macros -> eMailerOOo**.  
You may need to open the applications (Writer and Calc) and add the macros with the **Scope** set to the supported applications.

This only needs to be done once for LibreOffice and each application, and unfortunately I haven't found anything simpler yet.

___

## How to build the extension:

Normally, the extension is created with Eclipse for Java and [LOEclipse][74]. To work around Eclipse, I modified LOEclipse to allow the extension to be created with Apache Ant.  
To create the eMailerOOo extension with the help of Apache Ant, you need to:
- Install the [Java SDK][75] version 8 or higher.
- Install [Apache Ant][76] version 1.10.0 or higher.
- Install [LibreOffice and its SDK][77] version 7.x or higher.
- Clone the [eMailerOOo][78] repository on GitHub into a folder.
- From this folder, move to the directory: `source/eMailerOOo/`
- In this directory, edit the file: `build.properties` so that the `office.install.dir` and `sdk.dir` properties point to the folders where LibreOffice and its SDK were installed, respectively.
- Start the archive creation process using the command: `ant`
- You will find the generated archive in the subfolder: `dist/`

___

## Has been tested with:

* LibreOffice 7.3.7.2 - Lubuntu 22.04 - Python version 3.10.12 - OpenJDK-11-JRE (amd64)

* LibreOffice 7.5.4.2(x86) - Windows 10 - Python version 3.8.16 - Adoptium JDK Hotspot 11.0.19 (under Lubuntu 22.04 / VirtualBox 6.1.38)

* LibreOffice 7.4.3.2(x64) - Windows 10(x64) - Python version 3.8.15  - Adoptium JDK Hotspot 11.0.17 (x64) (under Lubuntu 22.04 / VirtualBox 6.1.38)

* LibreOffice 24.2.1.2 - Lubuntu 22.04

* LibreOffice 24.8.0.3 (x86_64) - Windows 10(x64) - Python version 3.9.19 (under Lubuntu 22.04 / VirtualBox 6.1.38)

* **Does not work with OpenOffice** see [bug 128569][79]. Having no solution, I encourage you to install **LibreOffice**.

I encourage you in case of problem :confused:  
to create an [issue][12]  
I will try to solve it :smile:

___

## Historical:

### [All changes are logged in the version History][80]

[1]: </img/emailer.svg#collapse>
[2]: <https://prrvchr.github.io/eMailerOOo/>
[3]: <https://prrvchr.github.io/eMailerOOo/README_fr>
[4]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/TermsOfUse_en>
[5]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/PrivacyPolicy_en>
[6]: <https://prrvchr.github.io/eMailerOOo/#what-has-been-done-for-version-152>
[7]: <https://prrvchr.github.io/>
[8]: <https://www.libreoffice.org/download/download-libreoffice/>
[9]: <https://www.openoffice.org/download/index.html>
[10]: <https://prrvchr.github.io/eMailerOOo/#sending-email-with-a-libreoffice-macro-in-basic>
[11]: <https://github.com/prrvchr/eMailerOOo>
[12]: <https://github.com/prrvchr/eMailerOOo/issues/new>
[13]: <https://github.com/sponsors/prrvchr>
[14]: <https://appdefensealliance.dev/casa>
[15]: <https://github.com/prrvchr/OAuth2OOo/blob/master/LOV_OAuth2OOo.pdf>
[16]: <https://img.shields.io/static/v1?label=Sponsor&message=%E2%9D%A4&logo=GitHub&color=%23fe8e86#right>
[17]: <https://prrvchr.github.io/OAuth2OOo/#requirement>
[18]: <https://prrvchr.github.io/jdbcDriverOOo/#requirement>
[19]: <https://prrvchr.github.io/OAuth2OOo/img/OAuth2OOo.svg#middle>
[20]: <https://prrvchr.github.io/OAuth2OOo/>
[21]: <https://github.com/prrvchr/OAuth2OOo/releases/latest/download/OAuth2OOo.oxt>
[22]: <https://img.shields.io/github/v/tag/prrvchr/OAuth2OOo?label=latest#right>
[23]: <https://prrvchr.github.io/jdbcDriverOOo/img/jdbcDriverOOo.svg#middle>
[24]: <https://prrvchr.github.io/jdbcDriverOOo/>
[25]: <https://github.com/prrvchr/jdbcDriverOOo/releases/latest/download/jdbcDriverOOo.oxt>
[26]: <https://img.shields.io/github/v/tag/prrvchr/jdbcDriverOOo?label=latest#right>
[27]: <https://prrvchr.github.io/vCardOOo/img/vCardOOo.svg#middle>
[28]: <https://prrvchr.github.io/vCardOOo/>
[29]: <https://github.com/prrvchr/vCardOOo/releases/latest/download/vCardOOo.oxt>
[30]: <https://img.shields.io/github/v/tag/prrvchr/vCardOOo?label=latest#right>
[31]: <https://en.wikipedia.org/wiki/Nextcloud>
[32]: <https://prrvchr.github.io/gContactOOo/img/gContactOOo.svg#middle>
[33]: <https://prrvchr.github.io/gContactOOo/>
[34]: <https://github.com/prrvchr/gContactOOo/releases/latest/download/gContactOOo.oxt>
[35]: <https://img.shields.io/github/v/tag/prrvchr/gContactOOo?label=latest#right>
[36]: <https://prrvchr.github.io/mContactOOo/img/mContactOOo.svg#middle>
[37]: <https://prrvchr.github.io/mContactOOo/>
[38]: <https://github.com/prrvchr/mContactOOo/releases/latest/download/mContactOOo.oxt>
[39]: <https://img.shields.io/github/v/tag/prrvchr/mContactOOo?label=latest#right>
[40]: <https://prrvchr.github.io/HyperSQLOOo/img/HyperSQLOOo.svg#middle>
[41]: <https://prrvchr.github.io/HyperSQLOOo/>
[42]: <https://github.com/prrvchr/HyperSQLOOo/releases/latest/download/HyperSQLOOo.oxt>
[43]: <https://img.shields.io/github/v/tag/prrvchr/HyperSQLOOo?label=latest#right>
[44]: <https://prrvchr.github.io/HyperSQLOOo/#how-to-import-data-from-a-calc-file>
[45]: <https://prrvchr.github.io/eMailerOOo/img/eMailerOOo.svg#middle>
[46]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/eMailerOOo.oxt>
[47]: <https://img.shields.io/github/downloads/prrvchr/eMailerOOo/latest/total?label=v1.5.1#right>
[48]: <./setup/>
[49]: <https://prrvchr.github.io/eMailerOOo/#merge-emails-with-mailing-lists>
[50]: <https://prrvchr.github.io/eMailerOOo/#configure-connection>
[51]: <https://prrvchr.github.io/eMailerOOo/#outgoing-emails>
[52]: <img/eMailerOOo-Merger1.png>
[53]: <img/eMailerOOo-Merger2.png>
[54]: <img/eMailerOOo-Merger3.png>
[55]: <img/eMailerOOo-Merger4.png>
[56]: <https://prrvchr.github.io/eMailerOOo/#available-recipients>
[57]: <img/eMailerOOo-Merger5.png>
[58]: <img/eMailerOOo-Merger6.png>
[59]: <img/eMailerOOo-Merger7.png>
[60]: <img/eMailerOOo-Merger8.png>
[61]: <img/eMailerOOo-Merger9.png>
[62]: <img/eMailerOOo-Merger10.png>
[63]: <img/eMailerOOo-Merger11.png>
[64]: <img/eMailerOOo-Ispdb1.png>
[65]: <img/eMailerOOo-Ispdb2.png>
[66]: <img/eMailerOOo-Ispdb3.png>
[67]: <img/eMailerOOo-Ispdb4.png>
[68]: <img/eMailerOOo-Ispdb5.png>
[69]: <img/eMailerOOo-Ispdb6.png>
[70]: <img/eMailerOOo-Spooler1.png>
[71]: <https://prrvchr.github.io/eMailerOOo/#spooler-activity-log>
[72]: <img/eMailerOOo-Spooler2.png>
[73]: <img/eMailerOOo-Spooler3.png>
[74]: <https://github.com/LibreOffice/loeclipse>
[75]: <https://adoptium.net/temurin/releases/?version=8&package=jdk>
[76]: <https://ant.apache.org/manual/install.html>
[77]: <https://downloadarchive.documentfoundation.org/libreoffice/old/7.6.7.2/>
[78]: <https://github.com/prrvchr/eMailerOOo.git>
[79]: <https://bz.apache.org/ooo/show_bug.cgi?id=128569>
[80]: <./change/>
