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

If you are using LibreOffice on Linux, then you are subject to [bug 139538](https://bugs.documentfoundation.org/show_bug.cgi?id=139538).  
To work around the problem, please uninstall the packages:
- libreoffice-sdbc-hsqldb
- libhsqldb1.8.0-java

If you still want to use the Embedded HsqlDB functionality provided by LibreOffice, then install the [HsqlDBembeddedOOo](https://prrvchr.github.io/HsqlDBembeddedOOo/) extension.  
OpenOffice and LibreOffice on Windows are not subject to this malfunction.

## Installation:

It seems important that the file was not renamed when it was downloaded.
If necessary, rename it before installing it.

- Install [OAuth2OOo.oxt](https://github.com/prrvchr/OAuth2OOo/raw/master/OAuth2OOo.oxt) extention version 0.0.5.

You must first install this extension, if it is not already installed.

- Install [smtpServerOOo.oxt](https://github.com/prrvchr/smtpServerOOo/raw/main/smtpServerOOo.oxt) extension version 0.0.1.

Restart LibreOffice / OpenOffice after installation.

## Historical:

What remains to be done:

- Rewriting of mailmerge.py (to be compatible with: SSL and StartTLS, OAuth2 authentication... ie: with Mozilla IspBD technology)
- Write an Wizard using Mozilla IspDB technology able to find the correct configuration working with mailmerge.py.
- Writing of a UNO Service, running in the background (Python Thread), allowing to send e-mails.
