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

# version [1.2.0][6]

## Introduction:

**eMailerOOo** is part of a [Suite][7] of [LibreOffice][8] ~~and/or [OpenOffice][9]~~ extensions allowing to offer you innovative services in these office suites.

This extension allows you to send documents in LibreOffice as an email, possibly by mail merge, to your telephone contacts.

It also provides an **API allowing you to send emails in BASIC** and supporting the most advanced technologies (OAuth2 protocol, Mozilla IspDB, HTTP instead of SMTP/IMAP, ...). A macro [SendEmail][10] for sending emails is provided as an example.  
If you open a document beforehand, you can launch it by:  
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

        This extension is only necessary if you want to use your contacts present on a [**Nextcloud**][29] platform as a data source for mailing lists and document merging.

    - [![gContactOOo logo][30]][31] Install **[gContactOOo.oxt][32]** extension [![Version][33]][32]

        This extension is only needed if you want to use your personal phone contacts (Android contact) as a data source for mailing lists and document merging.

    - [![mContactOOo logo][34]][35] Install **[mContactOOo.oxt][36]** extension [![Version][37]][36]

        This extension is only needed if you want to use your Microsoft Outlook contacts as a data source for mailing lists and document merging.

- ![eMailerOOo logo][38] Install **[eMailerOOo.oxt][39]** extension [![Version][40]][39]

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
- [vCardOOo][26]. This extension will allow you to use your contacts present on a [**Nextcloud**][29] platform as a data source.
- [gContactOOo][31]. This extension will allow you to use your Android phone (your phone contacts) as a datasource.
- [mContactOOo][35]. This extension will allow you to use your Microsoft Outlook contacts as a datasource.

For these 3 extensions the name of the **main table** can be found (and even changed before any connection) in:  
**Tools -> Options -> Internet -> Extension name -> Main table name**

This mode of use is made up of 3 sections:
- [Merge emails with mailing lists][41].
- [Configure connection][42].
- [Outgoing emails][43].

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

![eMailerOOo Merger screenshot 1][44]

#### Data source selection:

The datasource load for the **Email merging** wizard should appear: 

![eMailerOOo Merger screenshot 2][45]

The following screenshots use the [gContactOOo][31] extension as the **datasource**. If you are using your own **datasource**, it is necessary to adapt the settings in relation to it.

In the following screenshot, we can see that the **datasource** gContactOOo is called: `Addresses` and that in the list of tables the table: `PUBLIC.All my contacts` is selected.

![eMailerOOo Merger screenshot 3][46]

If no mailing list exists, you need to create one, by entering its name and validating with: `ENTER` or the `Add` button.

Make sure when creating the mailing list that the **main table** is always selected in the list of tables.  
If this recommendation is not followed then **merging of documents will not work** and this silently.

![eMailerOOo Merger screenshot 4][47]

Now that your new mailing list is available in the list, you need to select it.

And add the following columns:
- Primary key column: `Uri`
- Email address columns: `HomeEmail`, `WorkEmail` and `OtherEmail`

If several columns of email addresses are selected, then the order becomes relevant since the email will be sent to the first available address.  
In addition, on Recipients selection step of the wizard, in the [Available recipients][48] tab, only records with at least one email address column entered will be listed.  
So make sure you have an address book with at least one of the email address field (Home, Work or Other) entered.

![eMailerOOo Merger screenshot 5][49]

This setting is to be made only for new mailing lists.  
You can now proceed to the next step.

#### Recipients selection:

##### Available recipients:

The recipients are selected using 2 buttons `Add all` and `Add` allowing respectively:
- Either add the group of recipients selected from the `Address book` list. This allows during a mailing, that the modifications of the contents of the group are taken into account. A mailing list only accepts one group.
- Either add the selection, which can be multiple using the `CTRL` key. This selection is immutable regardless of the modification of the address book groups.

![eMailerOOo Merger screenshot 6][50]

Example of multiple selection:

![eMailerOOo Merger screenshot 7][51]

##### Selected recipients:

The recipients are deselected using 2 buttons `Remove all` and `Remove` allowing respectively:
- Either remove the group that has been assigned to this mailing list. This is necessary in order to be able to edit the content of this mailing list again.
- Either remove the selection, which can be multiple using the `CTRL` key. 

![eMailerOOo Merger screenshot 8][52]

If you have selected at least 1 recipient, you can proceed to the next step.

#### Sending options selection:

If this is not already done, you must create a new sender using the `Add` button.

![eMailerOOo Merger screenshot 9][53]

The creation of the new sender is described in the [Configure connection][42] section.

The email must have a subject. It can be saved in the Writer document.  
You can insert merge fields in the email subject. A merge field is composed of an opening brace, the name of the referenced column (case sensitive) and a closing brace (ie: `{ColumnName}`).

![eMailerOOo Merger screenshot 10][54]

The email may optionally have attached files. They can be saved in the Writer document.  
The following screenshot shows 1 attached file which will be merged on the data source then converted to PDF format before being attached to the email.

![eMailerOOo Merger screenshot 11][55]

Make sure to always exit the wizard with the `Finish` button to confirm submitting the send jobs.  
To submit mailing jobs, please follow the section [Outgoing emails][43].

### Configure connection:

#### Starting the connection wizard:

In LibreOffice / OpenOffice go to: **Tools -> Add-Ons -> Sending emails -> Configure connection**

![eMailerOOo Ispdb screenshot 1][56]

#### Account selection:

![eMailerOOo Ispdb screenshot 2][57]

#### Find the configuration:

![eMailerOOo Ispdb screenshot 3][58]

#### SMTP configuration:

![eMailerOOo Ispdb screenshot 4][59]

#### IMAP configuration:

![eMailerOOo Ispdb screenshot 5][60]

#### Connection test:

![eMailerOOo Ispdb screenshot 6][61]

Always exit the wizard with the `Finish` button to save the connection settings.

### Outgoing emails:

#### Starting the email spooler:

In LibreOffice / OpenOffice go to: **Tools -> Add-Ons -> Sending emails -> Outgoing emails**

![eMailerOOo Spooler screenshot 1][62]

#### List of outgoing emails:

Each send job has 3 different states:
- State **0**: the email is ready for sending.
- State **1**: the email was sent successfully.
- State **2**: An error occurred while sending the email. You can view the error message in the [Spooler activity log][63]. 

![eMailerOOo Spooler screenshot 2][64]

The email spooler is stopped by default. **It must be started with the `Start / Stop` button so that the pending emails are sent**.

#### Spooler activity log:

When the email spooler is started, its activity can be viewed in the activity log.

![eMailerOOo Spooler screenshot 3][65]

___

## Has been tested with:

* LibreOffice 7.3.7.2 - Lubuntu 22.04 - Python version 3.10.12 - OpenJDK-11-JRE (amd64)

* LibreOffice 7.5.4.2(x86) - Windows 10 - Python version 3.8.16 - Adoptium JDK Hotspot 11.0.19 (under Lubuntu 22.04 / VirtualBox 6.1.38)

* LibreOffice 7.4.3.2(x64) - Windows 10(x64) - Python version 3.8.15  - Adoptium JDK Hotspot 11.0.17 (x64) (under Lubuntu 22.04 / VirtualBox 6.1.38)

* **Does not work with OpenOffice on Windows** see [bug 128569][66]. Having no solution, I encourage you to install **LibreOffice**.

I encourage you in case of problem :confused:  
to create an [issue][11]  
I will try to solve it :smile:

___

## Historical:

### What has been done for version 0.0.1:

- Writing an [IspDB][67] or SMTP servers connection configuration wizard allowing:
    - Find the connection parameters to an SMTP server from an email address. Besides, I especially thank Mozilla, for [Thunderbird autoconfiguration database][68] or IspDB, which made this challenge possible...
    - Display the activity of the UNO service `com.sun.star.mail.MailServiceProvider` when connecting to the SMTP server and sending an email.

- Writing an email [Spooler][69] allowing:
    - View the email sending jobs with their respective status.
    - Display the activity of the UNO service `com.sun.star.mail.SpoolerService` when sending emails.
    - Start and stop the spooler service.

- Writing an email [Merger][70] allowing:
    - To create mailing lists.
    - To merge and convert the current document to HTML format to make it the email message.
    - To merge and/or convert in PDF format any possible files attached to the email. 

- Writing a document [Mailer][71] allowing:
    - To convert the document to HTML format to make it the email message.
    - To convert in PDF format any possible files attached to the email.

- Writing a [Grid][72] driven by a `com.sun.star.sdb.RowSet` allowing:
    - To be configurable on the columns to be displayed.
    - To be configurable on the sort order to be displayed.
    - Save the display settings.

### What has been done for version 0.0.2:

- Rewrite of [IspDB][67] or Mail servers connection configuration wizard in order to integrate the IMAP connection configuration.
    - Use of [IMAPClient][73] version 2.2.0: an easy-to-use, Pythonic and complete IMAP client library.
    - Extension of [com.sun.star.mail.*][74] IDL files:
        - [XMailMessage2.idl][75] now supports email threading.
        - The new [XImapService.idl][76] interface allows access to part of the IMAPClient library.

- Rewriting of the [Spooler][77] in order to integrate IMAP functionality such as the creation of a thread summarizing the mailing and grouping all the emails sent.

- Submitting the eMailerOOo extension to Google and obtaining permission to use its GMail API to send emails with a Google account.

### What has been done for version 0.0.3:

- Rewrote the [Grid][72] to allow:
    - Sorting on a column with the integration of the UNO service [SortableGridDataModel][78].
    - To generate the filter of records needed by the service [Spooler][69].
    - Sharing the python module with the [Grid][79] module of the [jdbcDriverOOo][22] extension.

- Rewrote the [Merger][70] to allow:
    - Schema name management in table names to be compatible with version 0.0.4 of [jdbcDriverOOo][22]
    - The creation of a mailing list on a group of the address book and allowing to follow the modification of its content.
    - The use of primary key, which can be composite, supporting [DataType][80] `VARCHAR` and `INTEGER` or derived.
    - A preview of the document with merge fields filled in faster thanks to the [Grid][72].

- Rewrote the [Spooler][69] to allow:
    - The use of new filters supporting composite primary keys provided by the [Merger][70].
    - The use of the new [Grid][72] allowing sorting on a column.

- Many other things...

### What has been done for version 1.0.0:

- The **smtpMailerOOo** extension has been renamed to **eMailerOOo**.

### What has been done for version 1.0.1:

- The absence or obsolescence of the **OAuth2OOo** and/or **jdbcDriverOOo** extensions necessary for the proper functioning of **eMailerOOo** now displays an error message. This is to prevent a malfunction such as [issue #3][81] from recurring...

- The underlying HsqlDB database can be opened in Base with: **Tools -> Options -> Internet -> eMailerOOo -> Database**.

- The **Tools -> Add-Ons** menu now displays correctly based on context.

- Many other things...

### What has been done for version 1.0.2:

- If no configuration is found in the connection configuration wizard (IspDB Wizard) then it is possible to configure the connection manually. See [issue #5][82].

### What has been done for version 1.1.0:

- In the connection configuration wizard (IspDB Wizard) it is now possible to deactivate the IMAP configuration.  
    As a result, this no longer sends a thread (IMAP message) when merging a mailing.  
    In this same wizard, it is now possible to enter an email reply-to address.

- In the email merge wizard, it is now possible to insert merge fields in the subject of the email. See [issue #6][83].  
    In the subject of an email, a merge field is composed of an opening brace, the name of the referenced column (case sensitive) and a closing brace (ie: `{ColumnName}`).  
    When entering the email subject, a syntax error in a merge field will be reported and will prevent the mailing from being submitted.

- It is now possible in the Spooler to view emails in eml format.

- A service [com.sun.star.mail.MailUser][84] now allows access to a connection configuration (SMTP and/or IMAP) from an email address following rfc822.  
    Another service [com.sun.star.datatransfer.TransferableFactory][85] allows, as its name suggests, the creation of [Transferable][86] from a String, a binary sequence, an Url (file://...) or a data stream (InputStream).  
    These two new services greatly simplify the LibreOffice mail API and allow sending emails from Basic. See [Issue #4][87].  
    You will find a Basic macro allowing you to send emails in: **Tools -> Macros -> Edit Macros... -> eMailerOOo -> SendEmail**.

### What has been done for version 1.1.1:

- Support for version **1.2.0** of the **OAuth2OOo** extension. Previous versions will not work with **OAuth2OOo** extension 1.2.0 or higher.

### What has been done for version 1.2.0:

- All Python packages necessary for the extension are now recorded in a [requirements.txt][88] file following [PEP 508][89].
- Now if you are not on Windows then the Python packages necessary for the extension can be easily installed with the command:  
  `pip install requirements.txt`
- Modification of the [Requirement][90] section.

### What remains to be done for version 1.2.0:

- Add new languages for internationalization...

- Anything welcome...

[1]: </img/emailer.svg#collapse>
[2]: <https://prrvchr.github.io/eMailerOOo/>
[3]: <https://prrvchr.github.io/eMailerOOo/README_fr>
[4]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/TermsOfUse_en>
[5]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/PrivacyPolicy_en>
[6]: <https://prrvchr.github.io/eMailerOOo/#what-has-been-done-for-version-110>
[7]: <https://prrvchr.github.io/>
[8]: <https://www.libreoffice.org/download/download-libreoffice/>
[9]: <https://www.openoffice.org/download/index.html>
[10]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/eMailerOOo/SendEmail.xba>
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
[38]: <https://prrvchr.github.io/eMailerOOo/img/eMailerOOo.svg#middle>
[39]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/eMailerOOo.oxt>
[40]: <https://img.shields.io/github/downloads/prrvchr/eMailerOOo/latest/total?label=v1.2.0#right>
[41]: <https://prrvchr.github.io/eMailerOOo/#merge-emails-with-mailing-lists>
[42]: <https://prrvchr.github.io/eMailerOOo/#configure-connection>
[43]: <https://prrvchr.github.io/eMailerOOo/#outgoing-emails>
[44]: <img/eMailerOOo-Merger1.png>
[45]: <img/eMailerOOo-Merger2.png>
[46]: <img/eMailerOOo-Merger3.png>
[47]: <img/eMailerOOo-Merger4.png>
[48]: <https://prrvchr.github.io/eMailerOOo/#available-recipients>
[49]: <img/eMailerOOo-Merger5.png>
[50]: <img/eMailerOOo-Merger6.png>
[51]: <img/eMailerOOo-Merger7.png>
[52]: <img/eMailerOOo-Merger8.png>
[53]: <img/eMailerOOo-Merger9.png>
[54]: <img/eMailerOOo-Merger10.png>
[55]: <img/eMailerOOo-Merger11.png>
[56]: <img/eMailerOOo-Ispdb1.png>
[57]: <img/eMailerOOo-Ispdb2.png>
[58]: <img/eMailerOOo-Ispdb3.png>
[59]: <img/eMailerOOo-Ispdb4.png>
[60]: <img/eMailerOOo-Ispdb5.png>
[61]: <img/eMailerOOo-Ispdb6.png>
[62]: <img/eMailerOOo-Spooler1.png>
[63]: <https://prrvchr.github.io/eMailerOOo/#spooler-activity-log>
[64]: <img/eMailerOOo-Spooler2.png>
[65]: <img/eMailerOOo-Spooler3.png>
[66]: <https://bz.apache.org/ooo/show_bug.cgi?id=128569>
[67]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/ispdb>
[68]: <https://wiki.mozilla.org/Thunderbird:Autoconfiguration>
[69]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler>
[70]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/merger>
[71]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/mailer>
[72]: <https://github.com/prrvchr/eMailerOOo/tree/master/uno/lib/uno/grid>
[73]: <https://github.com/mjs/imapclient#readme>
[74]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/idl/com/sun/star/mail>
[75]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailMessage2.idl>
[76]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XImapService.idl>
[77]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/spooler.py>
[78]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/grid/SortableGridDataModel.html>
[79]: <https://github.com/prrvchr/jdbcDriverOOo/tree/master/source/jdbcDriverOOo/service/pythonpath/jdbcdriver/grid>
[80]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/sdbc/DataType.html>
[81]: <https://github.com/prrvchr/eMailerOOo/issues/3>
[82]: <https://github.com/prrvchr/eMailerOOo/issues/5>
[83]: <https://github.com/prrvchr/eMailerOOo/issues/6>
[84]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailUser.idl>
[85]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/datatransfer/XTransferableFactory.idl>
[86]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/datatransfer/XTransferable.html>
[87]: <https://github.com/prrvchr/eMailerOOo/issues/4>
[88]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/requirements.txt>
[89]: <https://peps.python.org/pep-0508/>
[90]: <https://prrvchr.github.io/eMailerOOo/#requirement>
