[**Ce document en français**](https://prrvchr.github.io/smtpServerOOo/README_fr)

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

## Installation:

It seems important that the file was not renamed when it was downloaded.
If necessary, rename it before installing it.

- Install [OAuth2OOo.oxt](https://github.com/prrvchr/OAuth2OOo/raw/master/OAuth2OOo.oxt) extention version 0.0.5.

You must first install this extension, if it is not already installed.

- Install [smtpServerOOo.oxt](https://github.com/prrvchr/smtpServerOOo/raw/master/smtpServerOOo.oxt) extension version 0.0.1.

Restart LibreOffice / OpenOffice after installation.

## Historical:

What remains to be done:

- Rewriting of mailmerge.py (to be compatible with: SSL and StartTLS, OAuth2 authentication... ie: with Mozilla IspBD technology)
- Write an Wizard using Mozilla IspDB technology able to find the correct configuration working with mailmerge.py.
- Writing of a UNO Service, running in the background (Python Thread), allowing to send e-mails.
