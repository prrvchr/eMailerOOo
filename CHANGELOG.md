---
layout: default
title: eMailerOOo historical (English)
permalink: /change/
redirect_from:
  - /CHANGELOG
  - /CHANGELOG.html
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
# [![eMailerOOo logo][1]][2] Historical

**Ce [document][3] en français.**

Regarding installation, configuration and use, please consult the **[documentation][4]**.

### What was done for version 0.0.1:

- Writing an [IspDB][5] or SMTP servers connection configuration wizard allowing:
    - Find the connection parameters to an SMTP server from an email address. Besides, I especially thank Mozilla, for [Thunderbird autoconfiguration database][6] or IspDB, which made this challenge possible...
    - Display the activity of the UNO service `com.sun.star.mail.MailServiceProvider` when connecting to the SMTP server and sending an email.

- Writing an email [Spooler][7] allowing:
    - View the email sending jobs with their respective status.
    - Display the activity of the UNO service `com.sun.star.mail.SpoolerService` when sending emails.
    - Start and stop the spooler service.

- Writing an email [Merger][8] allowing:
    - To create mailing lists.
    - To merge and convert the current document to HTML format to make it the email message.
    - To merge and/or convert in PDF format any possible files attached to the email. 

- Writing a document [Mailer][9] allowing:
    - To convert the document to HTML format to make it the email message.
    - To convert in PDF format any possible files attached to the email.

- Writing a [Grid][10] driven by a `com.sun.star.sdb.RowSet` allowing:
    - To be configurable on the columns to be displayed.
    - To be configurable on the sort order to be displayed.
    - Save the display settings.

### What was done for version 0.0.2:

- Rewrite of [IspDB][5] or Mail servers connection configuration wizard in order to integrate the IMAP connection configuration.
    - Use of [IMAPClient][11] version 2.2.0: an easy-to-use, Pythonic and complete IMAP client library.
    - Extension of [com.sun.star.mail.*][12] IDL files:
        - [XMailMessage2.idl][13] now supports email threading.
        - The new [XImapService.idl][14] interface allows access to part of the IMAPClient library.

- Rewriting of the [Spooler][7] in order to integrate IMAP functionality such as the creation of a thread summarizing the mailing and grouping all the emails sent.

- Submitting the eMailerOOo extension to Google and obtaining permission to use its GMail API to send emails with a Google account.

### What was done for version 0.0.3:

- Rewrote the [Grid][10] to allow:
    - Sorting on a column with the integration of the UNO service [SortableGridDataModel][15].
    - To generate the filter of records needed by the service [Spooler][7].
    - Sharing the python module with the [Grid][16] module of the [jdbcDriverOOo][17] extension.

- Rewrote the [Merger][8] to allow:
    - Schema name management in table names to be compatible with version 0.0.4 of [jdbcDriverOOo][17]
    - The creation of a mailing list on a group of the address book and allowing to follow the modification of its content.
    - The use of primary key, which can be composite, supporting [DataType][18] `VARCHAR` and `INTEGER` or derived.
    - A preview of the document with merge fields filled in faster thanks to the [Grid][10].

- Rewrote the [Spooler][7] to allow:
    - The use of new filters supporting composite primary keys provided by the [Merger][8].
    - The use of the new [Grid][10] allowing sorting on a column.

- Many other things...

### What was done for version 1.0.0:

- The **smtpMailerOOo** extension has been renamed to **eMailerOOo**.

### What was done for version 1.0.1:

- The absence or obsolescence of the **OAuth2OOo** and/or **jdbcDriverOOo** extensions necessary for the proper functioning of **eMailerOOo** now displays an error message. This is to prevent a malfunction such as [issue #3][19] from recurring...

- The underlying HsqlDB database can be opened in Base with: **Tools -> Options -> Internet -> eMailerOOo -> Database**.

- The **Tools -> Add-Ons** menu now displays correctly based on context.

- Many other things...

### What was done for version 1.0.2:

- If no configuration is found in the connection configuration wizard (IspDB Wizard) then it is possible to configure the connection manually. See [issue #5][20].

### What was done for version 1.1.0:

- In the connection configuration wizard (IspDB Wizard) it is now possible to deactivate the IMAP configuration.  
    As a result, this no longer sends a thread (IMAP message) when merging a mailing.  
    In this same wizard, it is now possible to enter an email reply-to address.

- In the email merge wizard, it is now possible to insert merge fields in the subject of the email. See [issue #6][21].  
    In the subject of an email, a merge field is composed of an opening brace, the name of the referenced column (case sensitive) and a closing brace (ie: `{ColumnName}`).  
    When entering the email subject, a syntax error in a merge field will be reported and will prevent the mailing from being submitted.

- It is now possible in the Spooler to view emails in eml format.

- A service [com.sun.star.mail.MailUser][22] now allows access to a connection configuration (SMTP and/or IMAP) from an email address following rfc822.  
    Another service [com.sun.star.datatransfer.TransferableFactory][23] allows, as its name suggests, the creation of [Transferable][24] from a String, a binary sequence, an Url (file://...) or a data stream (InputStream).  
    These two new services greatly simplify the LibreOffice mail API and allow sending emails from Basic. See [Issue #4][25].  
    You will find a Basic macro allowing you to send emails in: **Tools -> Macros -> Edit Macros... -> eMailerOOo -> SendEmail**.

### What was done for version 1.1.1:

- Support for version **1.2.0** of the **OAuth2OOo** extension. Previous versions will not work with **OAuth2OOo** extension 1.2.0 or higher.

### What was done for version 1.2.0:

- All Python packages necessary for the extension are now recorded in a [requirements.txt][26] file following [PEP 508][27].
- Now if you are not on Windows then the Python packages necessary for the extension can be easily installed with the command:  
  `pip install requirements.txt`
- Modification of the [Requirement][28] section.

### What was done for version 1.2.1:

- Fixed a regression allowing errors to be displayed in the Spooler.
- Integration of a fix to workaround the [issue #159988][29].

### What was done for version 1.2.2:

- The creation of the database, during the first connection, uses the UNO API offered by the jdbcDriverOOo extension since version 1.3.2. This makes it possible to record all the information necessary for creating the database in 5 text tables which are in fact [5 csv files][30].
- The extension will ask you to install the OAuth2OOo and jdbcDriverOOo extensions in versions 1.3.4 and 1.3.2 respectively minimum.
- Many fixes.

### What was done for version 1.2.3:

- Fixed a regression from version 1.2.2 preventing jobs from being submitted to the email spooler.
- Fixed [issue #7][31] not allowing error messages to be displayed in case of incorrect configuration.

### What was done for version 1.2.4:

- Updated the [Python decorator][32] package to version 5.1.1.
- Updated the [Python ijson][33] package to version 3.3.0.
- Updated the [Python packaging][34] package to version 24.1.
- Updated the [Python setuptools][35] package to version 72.1.0 in order to respond to the [Dependabot security alert][36].
- Updated the [Python validators][37] package to version 0.33.0.
- The extension will ask you to install the OAuth2OOo and jdbcDriverOOo extensions in versions 1.3.6 and 1.4.2 respectively minimum.

### What was done for version 1.2.5:

- Updated the [Python setuptools][35] package to version 73.0.1.
- The extension will ask you to install the OAuth2OOo and jdbcDriverOOo extensions in versions 1.3.7 and 1.4.5 respectively minimum.
- Changes to extension options that require a restart of LibreOffice will result in a message being displayed.
- Support for LibreOffice version 24.8.x.

### What was done for version 1.2.6:

- If a reply address was given then it will be used when generating the eml file by the Spooler.
- The extension will ask you to install the OAuth2OOo and jdbcDriverOOo extensions in versions 1.3.8 and 1.4.6 respectively minimum.
- Modification of the extension options accessible via: **Tools -> Options... -> Internet -> eMailerOOo** in order to comply with the new graphic charter.

### What was done for version 1.2.7:

- The spooler allows opening sent emails either in the local email client (ie: Thunderbird) or online in your browser for accounts using an API for sending email (ie: Google and Microsoft).
- A new tab has been added to the spooler to allow tracking of mail service activity.
- Connections to Microsoft mail servers, which apparently no longer worked, have been migrated to the Graph API.
- For servers that no longer use SMTP and IMAP protocols and offer a replacement API (ie: Google API and Microsoft Graph):
    - All HTTP request parameters needed to send emails are stored in the LibreOffice configuration files.
    - All data needed to process HTTP responses are stored in the LibreOffice configuration files.

    This should allow implementing a third-party API for sending emails just by modifying the [Options.xcu][38] configuration file.
- To work, these new features require the OAuth2OOo extension in version 1.3.9 minimum.
- The command to open an email in Thunderbird can currently only be changed in the LibreOffice configuration (ie: Tools -> Options -> Advanced -> Open Expert Configuration).
- Non-refresh of scrollbars in multi-column lists (ie: grid) has been fixed and will be available from LibreOffice 24.8.4, see [SortableGridDataModel cannot be notified for changes][39].
- Opening emails in your browser does not work with a Microsoft account, the url allowing this has not yet been found and it seems that it would not be possible (ie: popup must be open by the Outlook window)?
- Many fixes.

### What was done for version 1.3.0:

- The extension will ask you to install the OAuth2OOo and jdbcDriverOOo extensions in versions 1.4.0 and 1.4.6 respectively minimum.
- Only providers with a third-party API or OAuth2 authentication and have an entry in the LibreOffice configuration will offer OAuth2 authentication by default in the connection setup wizard (ie: IspDB Wizard).
- The `yahoo.com` and `aol.com` email providers have been integrated. To make setup easier, a link to the page for creating an application password has been added to the connection setup wizard. If you think links to other providers are missing, please open an issue so I can add them back.
- Updated [Python IMAPClient][11] package to version 3.0.1.
- Thanks to improvements added to the [Eclipse plugin][40], it is now possible to create the extension file using the command line and the [Apache Ant][41] archive builder tool, see file [build.xml][42].
- The extension will refuse to install under OpenOffice regardless of version or LibreOffice other than 7.x or higher.
- Added binaries needed for Python libraries to work on Linux and LibreOffice 24.8 (ie: Python 3.9).
- Many fixes.

### What was done for version 1.3.1:

- Updated the [Python packaging][34] package to version 24.2.
- Updated the [Python setuptools][35] package to version 75.8.0.
- Updated the [Python six][43] package to version 1.17.0.
- Updated the [Python validators][37] package to version 0.34.0.
- Support for Python version 3.13.

### What was done for version 1.4.0:

- Updated the [Python packaging][34] package to version 25.0.
- Downgrade the [Python setuptools][35] package to version 75.3.2. to ensure support for Python 3.8.
- Passive registration deployment that allows for much faster installation of extensions and differentiation of registered UNO services from those provided by a Java or Python implementation. This passive registration is provided by the [LOEclipse][44] extension via [PR#152][45] and [PR#157][46].
- Modified [LOEclipse][44] to support the new `rdb` file format produced by the `unoidl-write` compilation utility. `idl` files have been updated to support both available compilation tools: idlc and unoidl-write.
- It is now possible to build the oxt file of the eMailerOOo extension only with the help of Apache Ant and a copy of the GitHub repository. The [How to build the extension][47] section has been added to the documentation.
- Implemented [PEP 570][48] in [logging][49] to support unique multiple arguments.
- To ensure the correct creation of the eMailerOOo database, it will be checked that the jdbcDriverOOo extension has `com.sun.star.sdb` as API level.
- Writing macros to be able to place custom menus wherever you want. To make it easier to create these custom menus, the section [How to customize LibreOffice menus][50] has been added to the documentation.
- Requires the **jdbcDriverOOo extension at least version 1.5.0**.
- Requires the **OAuth2OOo extension at least version 1.5.0**.

### What was done for version 1.4.1:

- In the connection wizard, if the given email address is not found in Mozilla IspDB or if you are offline, server names can be simple hostnames and valid ports will extend up to 65535. This is to address [issue#10][51].
- Fixed refresh issues on the second page of the connection wizard by using the UNO service `com.sun.star.awt.AsyncCallback`.
- Requires the **jdbcDriverOOo extension at least version 1.5.4**.
- Requires the **OAuth2OOo extension at least version 1.5.1**.

### What was done for version 1.4.2:

- Support for LibreOffice 25.2.x and 25.8.x on Windows 64-bit.
- Requires the **OAuth2OOo extension at least version 1.5.2**.

### What was done for version 1.5.0:

- Changed the wizard used when merging emails so that it opens in a dedicated window rather than modal as before.
- Also changed the email Spooler to open in a dedicated window rather than modal as before.
- These two new windows now display a progress bar as well as a status indicator when background tasks are launched.
- If tasks are started while these windows are requested to be closed, then the tasks will be canceled if possible and their completion will be waited for before closing.
- Completely rewritten the email Spooler, which now offers 3 tasks for sending emails, viewing an email, and merging a document:
  - [sender.py][52]
  - [mailer.py][53]
  - [viewer.py][54]
- Added the [XTaskEvent.idl][55] interface to the UNO API. This new interface, which is the transcription of the Python class [threading.Event][56], allows you to control a task executed by the LibreOffice [Dispatcher][57].
- If files are attached to an email and in PDF format, then they will follow the LibreOffice configuration settings found in: **File -> Export to -> Export as PDF** when being transformed.
- All methods needed for rendering and running in the background now use the UNO service [com.sun.star.awt.AsyncCallback][58] for callback.
- If the jdbcDriverOOo extension works without Java instrumentation, a warning message will be displayed in the extension options.
- Many corrections and some new features that I will let you discover.
- Requires the **jdbcDriverOOo extension at least version 1.6.0**.
- Requires the **OAuth2OOo extension at least version 1.6.0**.
- Has been tested with LibreOfficeDev 26.2.

### What was done for version 1.5.1:

- Any error occurring during the sending of an email will not affect its status if it happens during preparation and not during sending. This allows the error to be corrected and the sending attempt to be retried.
- It is possible to open an attachment directly from the list of files attached to an email.
- If this operation is performed on an attached file (Writer or Calc) to be merged, the merge fields of this document opened in LibreOffice will follow the selection of the Grids `Available recipients` and/or `Selected recipients`.
- All modal windows now open correctly in modal mode.
- Requires the **jdbcDriverOOo extension at least version 1.6.1**.
- Requires the **OAuth2OOo extension at least version 1.6.1**.

### What was done for version 1.5.2:

- Fixed a regression that prevented valid jobs from being submitted to the Spooler for merging.

### What was done for version 1.7.0:


### What remains to be done for version 1.7.0:

- Add new languages for internationalization...

- Anything welcome...

[1]: </img/emailer.svg#collapse>
[2]: <https://prrvchr.github.io/eMailerOOo/>
[3]: <https://prrvchr.github.io/eMailerOOo/change/fr/>
[4]: <https://prrvchr.github.io/eMailerOOo/>
[5]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/ispdb>
[6]: <https://wiki.mozilla.org/Thunderbird:Autoconfiguration>
[7]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler>
[8]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/merger>
[9]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/mailer>
[10]: <https://github.com/prrvchr/eMailerOOo/tree/master/uno/lib/uno/grid>
[11]: <https://github.com/mjs/imapclient#readme>
[12]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/idl/com/sun/star/mail>
[13]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailMessage2.idl>
[14]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XImapService.idl>
[15]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/grid/SortableGridDataModel.html>
[16]: <https://github.com/prrvchr/jdbcDriverOOo/tree/master/source/jdbcDriverOOo/service/pythonpath/jdbcdriver/grid>
[17]: <https://prrvchr.github.io/jdbcDriverOOo/README_fr>
[18]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/sdbc/DataType.html>
[19]: <https://github.com/prrvchr/eMailerOOo/issues/3>
[20]: <https://github.com/prrvchr/eMailerOOo/issues/5>
[21]: <https://github.com/prrvchr/eMailerOOo/issues/6>
[22]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailUser.idl>
[23]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/datatransfer/XTransferableFactory.idl>
[24]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/datatransfer/XTransferable.html>
[25]: <https://github.com/prrvchr/eMailerOOo/issues/4>
[26]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/requirements.txt>
[27]: <https://peps.python.org/pep-0508/>
[28]: <https://prrvchr.github.io/eMailerOOo/#requirement>
[29]: <https://bugs.documentfoundation.org/show_bug.cgi?id=159988>
[30]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/hsqldb>
[31]: <https://github.com/prrvchr/eMailerOOo/issues/7>
[32]: <https://pypi.org/project/decorator/>
[33]: <https://pypi.org/project/ijson/>
[34]: <https://pypi.org/project/packaging/>
[35]: <https://pypi.org/project/setuptools/>
[36]: <https://github.com/prrvchr/eMailerOOo/security/dependabot/1>
[37]: <https://pypi.org/project/validators/>
[38]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/Options.xcu>
[39]: <https://bugs.documentfoundation.org/show_bug.cgi?id=164040>
[40]: <https://github.com/LibreOffice/loeclipse/pull/123>
[41]: <https://ant.apache.org/>
[42]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/build.xml>
[43]: <https://pypi.org/project/six/>
[44]: <https://github.com/LibreOffice/loeclipse>
[45]: <https://github.com/LibreOffice/loeclipse/pull/152>
[46]: <https://github.com/LibreOffice/loeclipse/pull/157>
[47]: <https://prrvchr.github.io/eMailerOOo/#how-to-build-the-extension>
[48]: <https://peps.python.org/pep-0570/>
[49]: <https://github.com/prrvchr/eMailerOOo/blob/master/uno/lib/uno/logger/logwrapper.py#L106>
[50]: <https://prrvchr.github.io/eMailerOOo/#how-to-customize-libreoffice-menus>
[51]: <https://github.com/prrvchr/eMailerOOo/issues/10>
[52]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/thread/sender.py>
[53]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/thread/mailer.py>
[54]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/thread/viewer.py>
[55]: <https://github.com/prrvchr/eMailerOOo/blob/master/uno/rdb/idl/com/sun/star/task/XTaskEvent.idl>
[56]: <https://docs.python.org/3/library/threading.html#threading.Event>
[57]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XDispatch.html#dispatch>
[58]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/AsyncCallback.html>
