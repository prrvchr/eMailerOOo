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
# Documentation

**Ce [document][1] en français.**

**The use of this software subjects you to our [Terms Of Use][2] and [Data Protection Policy][3].**

# version [1.1.1][4]

## Introduction:

**eMailerOOo** is part of a [Suite][5] of [LibreOffice][6] ~~and/or [OpenOffice][7]~~ extensions allowing to offer you innovative services in these office suites.  

This extension allows you to send documents in LibreOffice as an email, possibly by mail merge, to your telephone contacts.  
It also provides an API usable in BASIC allowing you to send emails and supporting the most advanced technologies (OAuth2 protocol, Mozilla IspDB, HTTP instead of SMTP/IMAP, ...).  

Being free software I encourage you:
- To duplicate its [source code][8].
- To make changes, corrections, improvements.
- To open [issue][9] if needed.

In short, to participate in the development of this extension.  
Because it is together that we can make Free Software smarter.

___

## Requirement:

In order to take advantage of the latest versions of the Python libraries used in eMailerOOo, version 2 of Python has been abandoned in favor of **Python 3.8 minimum**.  
This means that **eMailerOOo no longer supports OpenOffice and LibreOffice 6.x on Windows since version 1.0.0**.
I can only advise you **to migrate to LibreOffice 7.x**.

eMailerOOo uses a local [HsqlDB][10] database version 2.7.2.  
HsqlDB being a database written in Java, its use requires the [installation and configuration][11] in LibreOffice / OpenOffice of a **JRE version 11 or later**.  
I recommend [Adoptium][12] as your Java installation source.

If you are using **LibreOffice Community on Linux**, you are subject to [bug 139538][13]. To work around the problem, please **uninstall the packages** with commands:
- `sudo apt remove libreoffice-sdbc-hsqldb` (to uninstall the libreoffice-sdbc-hsqldb package)
- `sudo apt remove libhsqldb1.8.0-java` (to uninstall the libhsqldb1.8.0-java package)

If you still want to use the Embedded HsqlDB functionality provided by LibreOffice, then install the [HyperSQLOOo][14] extension.  

___

## Installation:

It seems important that the file was not renamed when it was downloaded.
If necessary, rename it before installing it.

- Install ![OAuth2OOo logo][15] **[OAuth2OOo.oxt][16]** extension [![Version][17]][16]

    You must first install this extension, if it is not already installed.

- Install ![jdbcDriverOOo logo][18] **[jdbcDriverOOo.oxt][19]** extension [![Version][20]][19]

    This extension is necessary to use HsqlDB version 2.7.2 with all its features.

- If you don't have a datasource, you can install one of the following extensions:

    - ![vCardOOo logo][21] **[vCardOOo.oxt][22]** [![Version][23]][22]

        This extension is only necessary if you want to use your contacts present on a [**Nextcloud**][24] platform as a data source for mailing lists and document merging.

    - ![gContactOOo logo][25] **[gContactOOo.oxt][26]** [![Version][27]][26]

        This extension is only needed if you want to use your personal phone contacts (Android contact) as a data source for mailing lists and document merging.

    - ![mContactOOo logo][28] **[mContactOOo.oxt][29]** [![Version][30]][29]

        This extension is only needed if you want to use your Microsoft Outlook contacts as a data source for mailing lists and document merging.

- Install ![eMailerOOo logo][31] **[eMailerOOo.oxt][32]** extension [![Version][33]][32]

Restart LibreOffice / OpenOffice after installation.

___

## Use:

### Introduction:

To be able to use the email merge feature using mailing lists, it is necessary to have a **datasource** with tables having the following columns:
- One or more email addresses columns. These columns are chosen from a list and if this choice is not unique, then the first non-null email address column will be used.
- One or more primary key column to uniquely identify records, it can be a compound primary key. Supported types are VARCHAR and/or INTEGER, or derived. These columns must be declared with the NOT NULL constraint.

In addition, this **datasource** must have at least one **main table**, including all the records that can be used during the email merge.

If you do not have such a **datasource** then I invite you to install one of the following extensions:
- [vCardOOo][34]. This extension will allow you to use your contacts present on a [**Nextcloud**][24] platform as a data source.
- [gContactOOo][35]. This extension will allow you to use your Android phone (your phone contacts) as a datasource.
- [mContactOOo][36]. This extension will allow you to use your Microsoft Outlook contacts as a datasource.

For these 3 extensions the name of the **main table** can be found (and even changed before any connection) in:  
**Tools -> Options -> Internet -> Extension name -> Main table name**

This mode of use is made up of 3 sections:
- [Merge emails with mailing lists][37].
- [Configure connection][38].
- [Outgoing emails][39].

### Merge emails with mailing lists:

#### Requirement:

To be able to post emails to a mailing list, you must:
- Have a **datasource** as described in the previous introduction.
- Open a **new document** in LibreOffice / OpenOffice Writer.

This Writer document can include merge fields (insertable by the command: **Insert -> Field -> More fields -> Database -> Mail merge fields**), this is even necessary if you want to be able to customize the content of the email and any attached files.  
These merge fields should only refer to the **main table** of the **datasource**.

If you want to use an **existing Writer document**, you must also ensure that the **datasource** and the **main table** are attached to the document in: **Tools -> Address Book Source**.

If these recommendations are not followed then **merging of documents will not work** and this silently.

#### Starting the mail merge wizard:

In LibreOffice / OpenOffice Writer document go to: **Tools -> Add-Ons -> Sending emails -> Merge a document**

![eMailerOOo Merger screenshot 1][40]

#### Data source selection:

The datasource load for the **Email merging** wizard should appear: 

![eMailerOOo Merger screenshot 2][41]

The following screenshots use the [gContactOOo][35] extension as the **datasource**. If you are using your own **datasource**, it is necessary to adapt the settings in relation to it.

In the following screenshot, we can see that the **datasource** gContactOOo is called: `Addresses` and that in the list of tables the table: `PUBLIC.All my contacts` is selected.

![eMailerOOo Merger screenshot 3][42]

If no mailing list exists, you need to create one, by entering its name and validating with: `ENTER` or the `Add` button.

Make sure when creating the mailing list that the **main table** is always selected in the list of tables.  
If this recommendation is not followed then **merging of documents will not work** and this silently.

![eMailerOOo Merger screenshot 4][43]

Now that your new mailing list is available in the list, you need to select it.

And add the following columns:
- Primary key column: `Uri`
- Email address columns: `HomeEmail`, `WorkEmail` and `OtherEmail`

If several columns of email addresses are selected, then the order becomes relevant since the email will be sent to the first available address.  
In addition, on Recipients selection step of the wizard, in the [Available recipients][44] tab, only records with at least one email address column entered will be listed.  
So make sure you have an address book with at least one of the email address field (Home, Work or Other) entered.

![eMailerOOo Merger screenshot 5][45]

This setting is to be made only for new mailing lists.  
You can now proceed to the next step.

#### Recipients selection:

##### Available recipients:

The recipients are selected using 2 buttons `Add all` and `Add` allowing respectively:
- Either add the group of recipients selected from the `Address book` list. This allows during a mailing, that the modifications of the contents of the group are taken into account. A mailing list only accepts one group.
- Either add the selection, which can be multiple using the `CTRL` key. This selection is immutable regardless of the modification of the address book groups.

![eMailerOOo Merger screenshot 6][46]

Example of multiple selection:

![eMailerOOo Merger screenshot 7][47]

##### Selected recipients:

The recipients are deselected using 2 buttons `Remove all` and `Remove` allowing respectively:
- Either remove the group that has been assigned to this mailing list. This is necessary in order to be able to edit the content of this mailing list again.
- Either remove the selection, which can be multiple using the `CTRL` key. 

![eMailerOOo Merger screenshot 8][48]

If you have selected at least 1 recipient, you can proceed to the next step.

#### Sending options selection:

If this is not already done, you must create a new sender using the `Add` button.

![eMailerOOo Merger screenshot 9][49]

The creation of the new sender is described in the [Configure connection][38] section.

The email must have a subject. It can be saved in the Writer document.  
You can insert merge fields in the email subject. A merge field is composed of an opening brace, the name of the referenced column (case sensitive) and a closing brace (ie: `{ColumnName}`).

![eMailerOOo Merger screenshot 10][50]

The email may optionally have attached files. They can be saved in the Writer document.  
The following screenshot shows 1 attached file which will be merged on the data source then converted to PDF format before being attached to the email.

![eMailerOOo Merger screenshot 11][51]

Make sure to always exit the wizard with the `Finish` button to confirm submitting the send jobs.  
To submit mailing jobs, please follow the section [Outgoing emails][39].

### Configure connection:

#### Starting the connection wizard:

In LibreOffice / OpenOffice go to: **Tools -> Add-Ons -> Sending emails -> Configure connection**

![eMailerOOo Ispdb screenshot 1][52]

#### Account selection:

![eMailerOOo Ispdb screenshot 2][53]

#### Find the configuration:

![eMailerOOo Ispdb screenshot 3][54]

#### SMTP configuration:

![eMailerOOo Ispdb screenshot 4][55]

#### IMAP configuration:

![eMailerOOo Ispdb screenshot 5][56]

#### Connection test:

![eMailerOOo Ispdb screenshot 6][57]

Always exit the wizard with the `Finish` button to save the connection settings.

### Outgoing emails:

#### Starting the email spooler:

In LibreOffice / OpenOffice go to: **Tools -> Add-Ons -> Sending emails -> Outgoing emails**

![eMailerOOo Spooler screenshot 1][58]

#### List of outgoing emails:

Each send job has 3 different states:
- State **0**: the email is ready for sending.
- State **1**: the email was sent successfully.
- State **2**: An error occurred while sending the email. You can view the error message in the [Spooler activity log][59]. 

![eMailerOOo Spooler screenshot 2][60]

The email spooler is stopped by default. **It must be started with the `Start / Stop` button so that the pending emails are sent**.

#### Spooler activity log:

When the email spooler is started, its activity can be viewed in the activity log.

![eMailerOOo Spooler screenshot 3][61]

___

## Has been tested with:

* LibreOffice 7.3.7.2 - Lubuntu 22.04 - Python version 3.10.12 - OpenJDK-11-JRE (amd64)

* LibreOffice 7.5.4.2(x86) - Windows 10 - Python version 3.8.16 - Adoptium JDK Hotspot 11.0.19 (under Lubuntu 22.04 / VirtualBox 6.1.38)

* LibreOffice 7.4.3.2(x64) - Windows 10(x64) - Python version 3.8.15  - Adoptium JDK Hotspot 11.0.17 (x64) (under Lubuntu 22.04 / VirtualBox 6.1.38)

* **Does not work with OpenOffice on Windows** see [bug 128569][62]. Having no solution, I encourage you to install **LibreOffice**.

I encourage you in case of problem :confused:  
to create an [issue][9]  
I will try to solve it :smile:

___

## Historical:

### What has been done for version 0.0.1:

- Writing an [IspDB][63] or SMTP servers connection configuration wizard allowing:
    - Find the connection parameters to an SMTP server from an email address. Besides, I especially thank Mozilla, for [Thunderbird autoconfiguration database][64] or IspDB, which made this challenge possible...
    - Display the activity of the UNO service `com.sun.star.mail.MailServiceProvider` when connecting to the SMTP server and sending an email.

- Writing an email [Spooler][65] allowing:
    - View the email sending jobs with their respective status.
    - Display the activity of the UNO service `com.sun.star.mail.SpoolerService` when sending emails.
    - Start and stop the spooler service.

- Writing an email [Merger][66] allowing:
    - To create mailing lists.
    - To merge and convert the current document to HTML format to make it the email message.
    - To merge and/or convert in PDF format any possible files attached to the email. 

- Writing a document [Mailer][67] allowing:
    - To convert the document to HTML format to make it the email message.
    - To convert in PDF format any possible files attached to the email.

- Writing a [Grid][68] driven by a `com.sun.star.sdb.RowSet` allowing:
    - To be configurable on the columns to be displayed.
    - To be configurable on the sort order to be displayed.
    - Save the display settings.

### What has been done for version 0.0.2:

- Rewrite of [IspDB][63] or Mail servers connection configuration wizard in order to integrate the IMAP connection configuration.
    - Use of [IMAPClient][69] version 2.2.0: an easy-to-use, Pythonic and complete IMAP client library.
    - Extension of [com.sun.star.mail.*][70] IDL files:
        - [XMailMessage2.idl][71] now supports email threading.
        - The new [XImapService.idl][72] interface allows access to part of the IMAPClient library.

- Rewriting of the [Spooler][73] in order to integrate IMAP functionality such as the creation of a thread summarizing the mailing and grouping all the emails sent.

- Submitting the eMailerOOo extension to Google and obtaining permission to use its GMail API to send emails with a Google account.

### What has been done for version 0.0.3:

- Rewrote the [Grid][68] to allow:
    - Sorting on a column with the integration of the UNO service [SortableGridDataModel][74].
    - To generate the filter of records needed by the service [Spooler][65].
    - Sharing the python module with the [jdbcDriverOOo][75] extension.

- Rewrote the [Merger][66] to allow:
    - Schema name management in table names to be compatible with version 0.0.4 of [jdbcDriverOOo][76]
    - The creation of a mailing list on a group of the address book and allowing to follow the modification of its content.
    - The use of primary key, which can be composite, supporting [DataType][77] `VARCHAR` and `INTEGER` or derived.
    - A preview of the document with merge fields filled in faster thanks to the [Grid][68].

- Rewrote the [Spooler][65] to allow:
    - The use of new filters supporting composite primary keys provided by the [Merger][66].
    - The use of the new [Grid][68] allowing sorting on a column.

- Many other things...

### What has been done for version 1.0.0:

- The **smtpMailerOOo** extension has been renamed to **eMailerOOo**.

### What has been done for version 1.0.1:

- The absence or obsolescence of the **OAuth2OOo** and/or **jdbcDriverOOo** extensions necessary for the proper functioning of **eMailerOOo** now displays an error message. This is to prevent a malfunction such as [issue #3][78] from recurring...

- The underlying HsqlDB database can be opened in Base with: **Tools -> Options -> Internet -> eMailerOOo -> Database**.

- The **Tools -> Add-Ons** menu now displays correctly based on context.

- Many other things...

### What has been done for version 1.0.2:

- If no configuration is found in the connection configuration wizard (IspDB Wizard) then it is possible to configure the connection manually. See [issue #5][79].

### What has been done for version 1.1.0:

- In the connection configuration wizard (IspDB Wizard) it is now possible to deactivate the IMAP configuration.  
    As a result, this no longer sends a thread (IMAP message) when merging a mailing.  
    In this same wizard, it is now possible to enter an email reply-to address.

- In the email merge wizard, it is now possible to insert merge fields in the subject of the email. See [issue #6][80].  
    In the subject of an email, a merge field is composed of an opening brace, the name of the referenced column (case sensitive) and a closing brace (ie: `{ColumnName}`).  
    When entering the email subject, a syntax error in a merge field will be reported and will prevent the mailing from being submitted.

- It is now possible in the Spooler to view emails in eml format.

- A service [com.sun.star.mail.MailUser][81] now allows access to a connection configuration (SMTP and/or IMAP) from an email address following rfc822.  
    Another service [com.sun.star.datatransfer.TransferableFactory][82] allows, as its name suggests, the creation of [Transferable][83] from a String, a binary sequence, an Url (file://...) or a data stream (InputStream).  
    These two new services greatly simplify the LibreOffice mail API and allow sending emails from Basic. See [Issue #4][84].  
    You will find a Basic macro allowing you to send emails in: **Tools -> Macros -> Edit Macros... -> eMailerOOo -> SendEmail**.

### What has been done for version 1.1.1:

- Support for version **1.2.0** of the **OAuth2OOo** extension. Previous versions will not work with **OAuth2OOo** extension 1.2.0 or higher.

### What remains to be done for version 1.1.1:

- Add new languages for internationalization...

- Anything welcome...

[1]: <https://prrvchr.github.io/eMailerOOo/README_fr>
[2]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/TermsOfUse_en>
[3]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/PrivacyPolicy_en>
[4]: <https://prrvchr.github.io/eMailerOOo/#what-has-been-done-for-version-110>
[5]: <https://prrvchr.github.io/>
[6]: <https://www.libreoffice.org/download/download-libreoffice/>
[7]: <https://www.openoffice.org/download/index.html>
[8]: <https://github.com/prrvchr/eMailerOOo>
[9]: <https://github.com/prrvchr/eMailerOOo/issues/new>
[10]: <http://hsqldb.org/>
[11]: <https://wiki.documentfoundation.org/Documentation/HowTo/Install_the_correct_JRE_-_LibreOffice_on_Windows_10>
[12]: <https://adoptium.net/releases.html?variant=openjdk11>
[13]: <https://bugs.documentfoundation.org/show_bug.cgi?id=139538>
[14]: <https://prrvchr.github.io/HyperSQLOOo/>
[15]: <https://prrvchr.github.io/OAuth2OOo/img/OAuth2OOo.svg#middle>
[16]: <https://github.com/prrvchr/OAuth2OOo/releases/latest/download/OAuth2OOo.oxt>
[17]: <https://img.shields.io/github/downloads/prrvchr/OAuth2OOo/latest/total?label=v1.2.1#right>
[18]: <https://prrvchr.github.io/jdbcDriverOOo/img/jdbcDriverOOo.svg#middle>
[19]: <https://github.com/prrvchr/jdbcDriverOOo/releases/latest/download/jdbcDriverOOo.oxt>
[20]: <https://img.shields.io/github/downloads/prrvchr/jdbcDriverOOo/latest/total?label=v1.0.5#right>
[21]: <https://prrvchr.github.io/vCardOOo/img/vCardOOo.svg#middle>
[22]: <https://github.com/prrvchr/vCardOOo/releases/latest/download/vCardOOo.oxt>
[23]: <https://img.shields.io/github/downloads/prrvchr/vCardOOo/latest/total?label=v1.0.3#right>
[24]: <https://en.wikipedia.org/wiki/Nextcloud>
[25]: <https://prrvchr.github.io/gContactOOo/img/gContactOOo.svg#middle>
[26]: <https://github.com/prrvchr/gContactOOo/releases/latest/download/gContactOOo.oxt>
[27]: <https://img.shields.io/github/downloads/prrvchr/gContactOOo/latest/total?label=v1.0.3#right>
[28]: <https://prrvchr.github.io/mContactOOo/img/mContactOOo.svg#middle>
[29]: <https://github.com/prrvchr/mContactOOo/releases/latest/download/mContactOOo.oxt>
[30]: <https://img.shields.io/github/downloads/prrvchr/mContactOOo/latest/total?label=v1.0.3#right>
[31]: <https://prrvchr.github.io/eMailerOOo/img/eMailerOOo.svg#middle>
[32]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/eMailerOOo.oxt>
[33]: <https://img.shields.io/github/downloads/prrvchr/eMailerOOo/latest/total?label=v1.1.1#right>
[34]: <https://prrvchr.github.io/vCardOOo/>
[35]: <https://prrvchr.github.io/gContactOOo/>
[36]: <https://prrvchr.github.io/mContactOOo/>
[37]: <https://prrvchr.github.io/eMailerOOo/#merge-emails-with-mailing-lists>
[38]: <https://prrvchr.github.io/eMailerOOo/#configure-connection>
[39]: <https://prrvchr.github.io/eMailerOOo/#outgoing-emails>
[40]: <img/eMailerOOo-Merger1.png>
[41]: <img/eMailerOOo-Merger2.png>
[42]: <img/eMailerOOo-Merger3.png>
[43]: <img/eMailerOOo-Merger4.png>
[44]: <https://prrvchr.github.io/eMailerOOo/#available-recipients>
[45]: <img/eMailerOOo-Merger5.png>
[46]: <img/eMailerOOo-Merger6.png>
[47]: <img/eMailerOOo-Merger7.png>
[48]: <img/eMailerOOo-Merger8.png>
[49]: <img/eMailerOOo-Merger9.png>
[50]: <img/eMailerOOo-Merger10.png>
[51]: <img/eMailerOOo-Merger11.png>
[52]: <img/eMailerOOo-Ispdb1.png>
[53]: <img/eMailerOOo-Ispdb2.png>
[54]: <img/eMailerOOo-Ispdb3.png>
[55]: <img/eMailerOOo-Ispdb4.png>
[56]: <img/eMailerOOo-Ispdb5.png>
[57]: <img/eMailerOOo-Ispdb6.png>
[58]: <img/eMailerOOo-Spooler1.png>
[59]: <https://prrvchr.github.io/eMailerOOo/#spooler-activity-log>
[60]: <img/eMailerOOo-Spooler2.png>
[61]: <img/eMailerOOo-Spooler3.png>
[62]: <https://bz.apache.org/ooo/show_bug.cgi?id=128569>
[63]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/ispdb>
[64]: <https://wiki.mozilla.org/Thunderbird:Autoconfiguration>
[65]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler>
[66]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/merger>
[67]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/mailer>
[68]: <https://github.com/prrvchr/eMailerOOo/tree/master/uno/lib/uno/grid>
[69]: <https://github.com/mjs/imapclient#readme>
[70]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/idl/com/sun/star/mail>
[71]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailMessage2.idl>
[72]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XImapService.idl>
[73]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/spooler.py>
[74]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/grid/SortableGridDataModel.html>
[75]: <https://github.com/prrvchr/jdbcDriverOOo/tree/master/source/jdbcDriverOOo/service/pythonpath/jdbcdriver/grid>
[76]: <https://prrvchr.github.io/jdbcDriverOOo/>
[77]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/sdbc/DataType.html>
[78]: <https://github.com/prrvchr/eMailerOOo/issues/3>
[79]: <https://github.com/prrvchr/eMailerOOo/issues/5>
[80]: <https://github.com/prrvchr/eMailerOOo/issues/6>
[81]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailUser.idl>
[82]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/datatransfer/XTransferableFactory.idl>
[83]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/datatransfer/XTransferable.html>
[84]: <https://github.com/prrvchr/eMailerOOo/issues/4>
