# ![smtpMailerOOo logo][1] smtpMailerOOo

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

**Ce [document][2] en français.**

**The use of this software subjects you to our [Terms Of Use][3] and [Data Protection Policy][4]**

# version [0.0.3][5]

## Introduction:

**smtpMailerOOo** is part of a [Suite][6] of [LibreOffice][7] and/or [OpenOffice][8] extensions allowing to offer you innovative services in these office suites.  
This extension allows you to send documents in LibreOffice / OpenOffice as an email, possibly by mail merge, to your telephone contacts.

Being free software I encourage you:
- To duplicate its [source code][9].
- To make changes, corrections, improvements.
- To open [issue][10] if needed.

In short, to participate in the development of this extension.  
Because it is together that we can make Free Software smarter.

## Requirement:

smtpMailerOOo uses a local [HsqlDB][11] database version 2.5.1.  
HsqlDB being a database written in Java, its use requires the [installation and configuration][12] in LibreOffice / OpenOffice of a **JRE version 11 or later**.  
I recommend [Adoptium][13] as your Java installation source.

If you are using **LibreOffice on Linux**, then you are subject to [bug 139538][14].  
To work around the problem, please uninstall the packages:
- libreoffice-sdbc-hsqldb
- libhsqldb1.8.0-java

If you still want to use the Embedded HsqlDB functionality provided by LibreOffice, then install the [HsqlDBembeddedOOo][15] extension.  
OpenOffice and LibreOffice on Windows are not subject to this malfunction.

## Installation:

It seems important that the file was not renamed when it was downloaded.
If necessary, rename it before installing it.

- Install ![OAuth2OOo logo][16] **[OAuth2OOo.oxt][17]** extension version 0.0.6.  
You must first install this extension, if it is not already installed.

- Install ![jdbcDriverOOo logo][18] **[jdbcDriverOOo.oxt][19]** extension version 0.0.4.  
This extension is necessary to use HsqlDB version 2.5.1 with all its features.

- If you don't have a datasource, you can install one of the following extensions:

  - ![vCardOOo logo][20] **[vCardOOo.oxt][21]** version 0.0.1.  
  This extension is only necessary if you want to use your contacts present on a [**Nextcloud**][22] platform as a data source for mailing lists and document merging.

  - ![gContactOOo logo][23] **[gContactOOo.oxt][24]** version 0.0.6.  
  This extension is only needed if you want to use your personal phone contacts (Android contact) as a data source for mailing lists and document merging.

  - ![mContactOOo logo][25] **[mContactOOo.oxt][26]** version 0.0.1.  
  This extension is only needed if you want to use your Microsoft Outlook contacts as a data source for mailing lists and document merging.

- Install ![smtpMailerOOo logo][27] **[smtpMailerOOo.oxt][28]** extension version 0.0.3.  

Restart LibreOffice / OpenOffice after installation.

## Use:

### Introduction:

To be able to use the email merge feature using mailing lists, it is necessary to have a datasource with tables having the following columns:
- One or more email addresses columns. These columns are chosen from a list and if this choice is not unique, then the first non-null email address column will be used.
- One or more primary key column to uniquely identify records, it can be a compound primary key. Supported types are VARCHAR and/or INTEGER, or derived. These columns must be declared with the NOT NULL constraint.

In addition, this datasource must have at least one **main table**, including all the records that can be used during the email merge.

If you do not have such a data source then I invite you to install one of the following extensions:
- [vCardOOo][21] extension.  
This extension will allow you to use your contacts present on a [**Nextcloud**][22] platform as a data source.
- [gContactOOo][24] extension.  
This extension will allow you to use your Android phone (your phone contacts) as a datasource.
- [mContactOOo][25] extension.  
This extension will allow you to use your Microsoft Outlook contacts as a datasource.

This mode of use is made up of 3 sections:
- [Merging emails to mailing lists][29].
- [Configure connection][30].
- [Outgoing emails][31].

### Merging emails to mailing lists:

#### Requirement:

To be able to post emails to a mailing list, you must:
- Have a data source as described in the previous introduction.
- Open a Writer document in LibreOffice / OpenOffice.

This Writer document can include merge fields (insertable by the command: **Insert -> Field -> More fields -> Database -> Mail merge fields**), this is even necessary if you want to be able to customize the content of the email.  
These merge fields should only reference the **main table** of the data source.

#### Starting the mail merge wizard:

In LibreOffice / OpenOffice Writer document go to: **Tools -> Add-Ons -> Sending emails -> Merge a document**

![smtpMailerOOo Merger screenshot 1][32]

#### Data source selection:

The datasource load for the **Email merging** wizard should appear: 

![smtpMailerOOo Merger screenshot 2][33]

The following screenshots use the [gContactOOo][24] extension as the data source. If you are using your own data source, it is necessary to adapt the settings in relation to it.

In the following screenshot, we can see that the data source gContactOOo is called: `Addresses` and that the **main table**: `PUBLIC.All my contacts` is selected.

![smtpMailerOOo Merger screenshot 3][34]

If no mailing list exists, you need to create one, by entering its name and validating with: `ENTER` or the `Add` button.

Make sure when creating the mailing list that the **main table** is always selected.

![smtpMailerOOo Merger screenshot 4][35]

Now that your new mailing list is available in the list, you need to select it.

And add the following columns:
- Primary key column: `Resource`
- Email address columns: `HomeEmail`, `WorkEmail` and `OtherEmail`

If several columns of email addresses are selected, then the order becomes relevant since the email will be sent to the first available address.  
In addition, on Recipients selection step of the wizard, in the [Available recipients][36] tab, only records with at least one email address column entered will be listed.  
So make sure you have an address book with at least one of the email address columns entered.

![smtpMailerOOo Merger screenshot 5][37]

This setting is to be made only for new mailing lists.  
You can now proceed to the next step.

#### Recipients selection:

##### Available recipients:

The recipients are selected using 2 buttons `Add all` and `Add` allowing respectively:
- Either add the group of recipients selected from the `Address book` list. This allows during a mailing, that the modifications of the contents of the group are taken into account. A mailing list only accepts one group.
- Either add the selection, which can be multiple using the `CTRL` key. This selection is immutable regardless of the modification of the address book groups.

![smtpMailerOOo Merger screenshot 6][38]

Example of multiple selection:

![smtpMailerOOo Merger screenshot 7][39]

##### Selected recipients:

The recipients are deselected using 2 buttons `Remove all` and `Remove` allowing respectively:
- Either remove the group that has been assigned to this mailing list. This is necessary in order to be able to edit the content of this mailing list again.
- Either remove the selection, which can be multiple using the `CTRL` key. 

![smtpMailerOOo Merger screenshot 8][40]

If you have selected at least 1 recipient, you can proceed to the next step.

#### Sending options selection:

If this is not already done, you must create a new sender using the `Add` button.

![smtpMailerOOo Merger screenshot 9][41]

The creation of the new sender is described in the [Configure connection][42] section.

The email must have a subject. It can be saved in the Writer document.

![smtpMailerOOo Merger screenshot 10][43]

The email may optionally have attached files. They can be saved in the Writer document.  
The following screenshot shows 1 attached file which will be merged on the data source then converted to PDF format before being attached to the email.

![smtpMailerOOo Merger screenshot 11][44]

Make sure to always exit the wizard with the `Finish` button to confirm submitting the send jobs.  
To submit mailing jobs, please follow the section [Outgoing emails][45].

### Configure connection:

#### Starting the connection wizard:

In LibreOffice / OpenOffice go to: **Tools -> Add-Ons -> Sending emails -> Configure connection**

![smtpMailerOOo Ispdb screenshot 1][46]

#### Account selection:

![smtpMailerOOo Ispdb screenshot 2][47]

#### Find the configuration:

![smtpMailerOOo Ispdb screenshot 3][48]

#### SMTP configuration:

![smtpMailerOOo Ispdb screenshot 4][49]

#### IMAP configuration:

![smtpMailerOOo Ispdb screenshot 5][50]

#### Connection test:

![smtpMailerOOo Ispdb screenshot 6][51]

Always exit the wizard with the `Finish` button to save the connection settings.

### Outgoing emails:

#### Starting the email spooler:

In LibreOffice / OpenOffice go to: **Tools -> Add-Ons -> Sending emails -> Outgoing emails**

![smtpMailerOOo Spooler screenshot 1][52]

#### List of outgoing emails:

Each send job has 3 different states:
- State **0**: the email is ready for sending.
- State **1**: the email was sent successfully.
- State **2**: An error occurred while sending the email. You can view the error message in the [Spooler activity log][53]. 

![smtpMailerOOo Spooler screenshot 2][54]

The email spooler is stopped by default. **It must be started with the `Start / Stop` button so that the pending emails are sent**.

#### Spooler activity log:

When the email spooler is started, its activity can be viewed in the activity log.

![smtpMailerOOo Spooler screenshot 3][55]

## Historical:

### What has been done for version 0.0.1:

- Writing an [IspDB][56] or SMTP servers connection configuration wizard allowing:
    - Find the connection parameters to an SMTP server from an email address. Besides, I especially thank Mozilla, for [Thunderbird autoconfiguration database][57] or IspDB, which made this challenge possible...
    - Display the activity of the UNO service `com.sun.star.mail.MailServiceProvider` when connecting to the SMTP server and sending an email.

- Writing an email [Spooler][58] allowing:
    - View the email sending jobs with their respective status.
    - Display the activity of the UNO service `com.sun.star.mail.SpoolerService` when sending emails.
    - Start and stop the spooler service.

- Writing an email [Merger][59] allowing:
    - To create mailing lists.
    - To merge and convert the current document to HTML format to make it the email message.
    - To merge and/or convert in PDF format any possible files attached to the email. 

- Writing a document [Mailer][60] allowing:
    - To convert the document to HTML format to make it the email message.
    - To convert in PDF format any possible files attached to the email.

- Writing a [Grid][61] driven by a `com.sun.star.sdb.RowSet` allowing:
    - To be configurable on the columns to be displayed.
    - To be configurable on the sort order to be displayed.
    - Save the display settings.

### What has been done for version 0.0.2:

- Rewrite of [IspDB][54] or Mail servers connection configuration wizard in order to integrate the IMAP connection configuration.
  - Use of [IMAPClient][62] version 2.2.0: an easy-to-use, Pythonic and complete IMAP client library.
  - Extension of [com.sun.star.mail.*][63] IDL files:
    - [XMailMessage2.idl][64] now supports email threading.
    - The new [XImapService.idl][65] interface allows access to part of the IMAPClient library.

- Rewriting of the [Spooler][66] in order to integrate IMAP functionality such as the creation of a thread summarizing the mailing and grouping all the emails sent.

- Submitting the smtpMailerOOo extension to Google and obtaining permission to use its GMail API to send emails with a Google account.

### What has been done for version 0.0.3:

- Rewrote the [Grid][61] to allow:
  - Sorting on a column with the integration of the UNO service [SortableGridDataModel][67].
  - To generate the filter of records needed by the service [Spooler][58].
  - Sharing the python module with the [jdbcDriverOOo][68] extension.

- Rewrote the [Merger][59] to allow:
  - Schema name management in table names to be compatible with version 0.0.4 of [jdbcDriverOOo][69]
  - The creation of a mailing list on a group of the address book and allowing to follow the modification of its content.
  - The use of primary key, which can be composite, supporting [DataType][70] `VARCHAR` and `INTEGER` or derived.
  - A preview of the document with merge fields filled in faster thanks to the [Grid][61].

- Rewrote the [Spooler][58] to allow:
  - The use of new filters supporting composite primary keys provided by the [Merger][59].
  - The use of the new [Grid][61] allowing sorting on a column.

- Many other things...

### What remains to be done for version 0.0.3:

- Add new languages for internationalization...

- Anything welcome...

[1]: <https://prrvchr.github.io/smtpMailerOOo/img/smtpMailerOOo.png>
[2]: <https://prrvchr.github.io/smtpMailerOOo/README_fr>
[3]: <https://prrvchr.github.io/smtpMailerOOo/source/smtpMailerOOo/registration/TermsOfUse_en>
[4]: <https://prrvchr.github.io/smtpMailerOOo/source/smtpMailerOOo/registration/PrivacyPolicy_en>
[5]: <https://prrvchr.github.io/smtpMailerOOo/#what-has-been-done-for-version-003>
[6]: <https://prrvchr.github.io/>
[7]: <https://www.libreoffice.org/download/download-libreoffice/>
[8]: <https://www.openoffice.org/download/index.html>
[9]: <https://github.com/prrvchr/smtpMailerOOo>
[10]: <https://github.com/prrvchr/smtpMailerOOo/issues/new>
[11]: <http://hsqldb.org/>
[12]: <https://wiki.documentfoundation.org/Documentation/HowTo/Install_the_correct_JRE_-_LibreOffice_on_Windows_10>
[13]: <https://adoptium.net/releases.html?variant=openjdk11>
[14]: <https://bugs.documentfoundation.org/show_bug.cgi?id=139538>
[15]: <https://prrvchr.github.io/HsqlDBembeddedOOo/>
[16]: <https://prrvchr.github.io/OAuth2OOo/img/OAuth2OOo.png>
[17]: <https://github.com/prrvchr/OAuth2OOo/raw/master/OAuth2OOo.oxt>
[18]: <https://prrvchr.github.io/jdbcDriverOOo/img/jdbcDriverOOo.png>
[19]: <https://github.com/prrvchr/jdbcDriverOOo/raw/master/jdbcDriverOOo.oxt>
[20]: <https://prrvchr.github.io/vCardOOo/img/vCardOOo.png>
[21]: <https://github.com/prrvchr/vCardOOo/raw/main/vCardOOo.oxt>
[22]: <https://en.wikipedia.org/wiki/Nextcloud>
[23]: <https://prrvchr.github.io/gContactOOo/img/gContactOOo.png>
[24]: <https://github.com/prrvchr/gContactOOo/raw/master/gContactOOo.oxt>
[25]: <https://prrvchr.github.io/mContactOOo/img/mContactOOo.png>
[26]: <https://github.com/prrvchr/mContactOOo/raw/main/mContactOOo.oxt>
[27]: <img/smtpMailerOOo.png>
[28]: <https://github.com/prrvchr/smtpMailerOOo/raw/master/smtpMailerOOo.oxt>
[29]: <https://prrvchr.github.io/smtpMailerOOo/#merging-emails-to-mailing-lists>
[30]: <https://prrvchr.github.io/smtpMailerOOo/#configure-connection>
[31]: <https://prrvchr.github.io/smtpMailerOOo/#outgoing-emails>
[32]: <img/smtpMailerOOo-Merger1.png>
[33]: <img/smtpMailerOOo-Merger2.png>
[34]: <img/smtpMailerOOo-Merger3.png>
[35]: <img/smtpMailerOOo-Merger4.png>
[36]: <https://prrvchr.github.io/smtpMailerOOo/#available-recipients>
[37]: <img/smtpMailerOOo-Merger5.png>
[38]: <img/smtpMailerOOo-Merger6.png>
[39]: <img/smtpMailerOOo-Merger7.png>
[40]: <img/smtpMailerOOo-Merger8.png>
[41]: <img/smtpMailerOOo-Merger9.png>
[42]: <https://prrvchr.github.io/smtpMailerOOo/#configure-connection>
[43]: <img/smtpMailerOOo-Merger10.png>
[44]: <img/smtpMailerOOo-Merger11.png>
[45]: <https://prrvchr.github.io/smtpMailerOOo/#outgoing-emails>
[46]: <img/smtpMailerOOo-Ispdb1.png>
[47]: <img/smtpMailerOOo-Ispdb2.png>
[48]: <img/smtpMailerOOo-Ispdb3.png>
[49]: <img/smtpMailerOOo-Ispdb4.png>
[50]: <img/smtpMailerOOo-Ispdb5.png>
[51]: <img/smtpMailerOOo-Ispdb6.png>
[52]: <img/smtpMailerOOo-Spooler1.png>
[53]: <https://prrvchr.github.io/smtpMailerOOo/#spooler-activity-log>
[54]: <img/smtpMailerOOo-Spooler2.png>
[55]: <img/smtpMailerOOo-Spooler3.png>
[56]: <https://github.com/prrvchr/smtpMailerOOo/tree/master/source/smtpMailerOOo/service/pythonpath/smtpmailer/ispdb>
[57]: <https://wiki.mozilla.org/Thunderbird/ISPDB>
[58]: <https://github.com/prrvchr/smtpMailerOOo/tree/master/source/smtpMailerOOo/service/pythonpath/smtpmailer/spooler>
[59]: <https://github.com/prrvchr/smtpMailerOOo/tree/master/source/smtpMailerOOo/service/pythonpath/smtpmailer/merger>
[60]: <https://github.com/prrvchr/smtpMailerOOo/tree/master/source/smtpMailerOOo/service/pythonpath/smtpmailer/mailer>
[61]: <https://github.com/prrvchr/smtpMailerOOo/tree/master/uno/lib/uno/grid>
[62]: <https://github.com/mjs/imapclient#readme>
[63]: <https://github.com/prrvchr/smtpMailerOOo/tree/master/source/smtpMailerOOo/idl/com/sun/star/mail>
[64]: <https://github.com/prrvchr/smtpMailerOOo/blob/master/source/smtpMailerOOo/idl/com/sun/star/mail/XMailMessage2.idl>
[65]: <https://github.com/prrvchr/smtpMailerOOo/blob/master/source/smtpMailerOOo/idl/com/sun/star/mail/XImapService.idl>
[66]: <https://github.com/prrvchr/smtpMailerOOo/tree/master/source/smtpMailerOOo/service/pythonpath/smtpmailer/mailspooler.py>
[67]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/grid/SortableGridDataModel.html>
[68]: <https://github.com/prrvchr/jdbcDriverOOo/tree/master/source/jdbcDriverOOo/service/pythonpath/jdbcdriver/grid>
[69]: <https://prrvchr.github.io/jdbcDriverOOo/>
[70]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/sdbc/DataType.html>
