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
**Ce [document](https://prrvchr.github.io/smtpServerOOo/README_fr) en français.**

**The use of this software subjects you to our** [**Terms Of Use**](https://prrvchr.github.io/smtpServerOOo/smtpServerOOo/registration/TermsOfUse_en) **and** [**Data Protection Policy**](https://prrvchr.github.io/smtpServerOOo/smtpServerOOo/registration/PrivacyPolicy_en)

**This extension is under development and is not yet available ... Thank you for your patience.**

# version [0.0.1](https://prrvchr.github.io/smtpServerOOo#historical)

## Introduction:

**smtpServerOOo** is part of a [Suite](https://prrvchr.github.io/) of [LibreOffice](https://fr.libreoffice.org/download/telecharger-libreoffice/) and/or [OpenOffice](https://www.openoffice.org/fr/Telecharger/) extensions allowing to offer you innovative services in these office suites.  
This extension allows you to send electronic mail in LibreOffice / OpenOffice, by a new smtp Client who act like a server.

Being free software I encourage you:
- To duplicate its [source code](https://github.com/prrvchr/smtpServerOOo).
- To make changes, corrections, improvements.
- To open [issue](https://github.com/prrvchr/smtpServerOOo/issues/new) if needed.

In short, to participate in the development of this extension.
Because it is together that we can make Free Software smarter.

## Requirement:

smtpServerOOo uses a local HsqlDB database of version 2.5.1.  
The use of HsqlDB requires the installation and configuration within LibreOffice / OpenOffice of a **JRE version 1.8 minimum** (ie: Java version 8)  
I recommend [AdoptOpenJDK](https://adoptopenjdk.net/) as your Java installation source.

If you are using **LibreOffice on Linux**, then you are subject to [bug 139538](https://bugs.documentfoundation.org/show_bug.cgi?id=139538).  
To work around the problem, please uninstall the packages:
- libreoffice-sdbc-hsqldb
- libhsqldb1.8.0-java

If you still want to use the Embedded HsqlDB functionality provided by LibreOffice, then install the [HsqlDBembeddedOOo](https://prrvchr.github.io/HsqlDBembeddedOOo/) extension.  
OpenOffice and LibreOffice on Windows are not subject to this malfunction.

## Installation:

It seems important that the file was not renamed when it was downloaded.
If necessary, rename it before installing it.

- Install [OAuth2OOo.oxt](https://github.com/prrvchr/OAuth2OOo/raw/master/OAuth2OOo.oxt) extension version 0.0.5.

You must first install this extension, if it is not already installed.

- Install [HsqlDBDriverOOo.oxt](https://github.com/prrvchr/HsqlDBDriverOOo/raw/master/HsqlDBDriverOOo.oxt) extension version 0.0.4.

This extension is necessary to use HsqlDB version 2.5.1 with all its features.

- Install [smtpServerOOo.oxt](https://github.com/prrvchr/smtpServerOOo/raw/main/smtpServerOOo.oxt) extension version 0.0.1.

Restart LibreOffice / OpenOffice after installation.

## Historical:

What remains to be done:

- Rewriting of mailmerge.py (to be compatible with: SSL and StartTLS, OAuth2 authentication... ie: with Mozilla IspBD technology)
- Write an Wizard using Mozilla IspDB technology able to find the correct configuration working with mailmerge.py.
- Writing of a UNO Service, running in the background (Python Thread), allowing to send e-mails.
