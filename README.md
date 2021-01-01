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
The use of HsqlDB requires the installation and configuration within  
LibreOffice / OpenOffice of a **JRE version 1.8 minimum** (ie: Java version 8)

Sometimes it may be necessary for LibreOffice users must have no HsqlDB driver installed with LibreOffice  
(check your Installed Application under Windows or your Packet Manager under Linux).  
~~It seems that version 7.x of LibreOffice has fixed this problem and is able to work with different driver version of HsqlDB simultaneously.~~  
After much testing it seems that LibreOffice (6.4.x and 7.x) cannot load a provided HsqlDB driver (hsqldb.jar v2.5.1), if the Embedded HsqlDB driver is installed (and even the solution is sometimes to rename the hsqldb.jar in /usr/share/java, uninstalling the libreoffice-sdbc-hsqldb package does not seem sufficient...)  
To overcome this limitation and if you want to use build-in Embedded HsqlDB, remove the build-in Embedded HsqlDB driver (hsqldb.jar v1.8.0) and install this extension: [HsqlDBembeddedOOo](https://prrvchr.github.io/HsqlDBembeddedOOo/) to replace the failing LibreOffice Embedded HsqlDB built-in driver.  
OpenOffice doesn't seem to need this workaround.

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
