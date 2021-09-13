<<<<<<< HEAD
**This [document](https://prrvchr.github.io/smtpMailerOOo) in English.**

**L'utilisation de ce logiciel vous soumet à nos** [**Conditions d'utilisation**](https://prrvchr.github.io/smtpMailerOOo/smtpMailerOOo/registration/TermsOfUse_fr)

**Cette extension est en cours de développement et n'est pas encore disponible ... Merci de votre patience.**

# version [0.0.1](https://prrvchr.github.io/smtpMailerOOo/README_fr#historique)

## Introduction:

**smtpMailerOOo** fait partie d'une [Suite](https://prrvchr.github.io/README_fr) d'extensions [LibreOffice](https://fr.libreoffice.org/download/telecharger-libreoffice/) et/ou [OpenOffice](https://www.openoffice.org/fr/Telecharger/) permettant de vous offrir des services inovants dans ces suites bureautique.  
Cette extension vous permet d'envoyer du courrier électronique dans LibreOffice / OpenOffice, éventuellement par publipostage, à vos contacts téléphoniques.

Etant un logiciel libre je vous encourage:
- A dupliquer son [code source](https://github.com/prrvchr/smtpMailerOOo).
- A apporter des modifications, des corrections, des ameliorations.
- D'ouvrir un [dysfonctionnement](https://github.com/prrvchr/smtpMailerOOo/issues/new) si nécessaire.
=======
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
**This [document](https://prrvchr.github.io/smtpServerOOo) in English.**

**L'utilisation de ce logiciel vous soumet à nos** [**Conditions d'utilisation**](https://prrvchr.github.io/smtpServerOOo/smtpServerOOo/registration/TermsOfUse_fr) **et à notre** [**Politique de protection des données**](https://prrvchr.github.io/smtpServerOOo/smtpServerOOo/registration/PrivacyPolicy_fr)

**Cette extension est en cours de développement et n'est pas encore disponible ... Merci de votre patience.**

# version [0.0.1](https://prrvchr.github.io/smtpServerOOo/README_fr#historique)

## Introduction:

**smtpServerOOo** fait partie d'une [Suite](https://prrvchr.github.io/README_fr) d'extensions [LibreOffice](https://fr.libreoffice.org/download/telecharger-libreoffice/) et/ou [OpenOffice](https://www.openoffice.org/fr/Telecharger/) permettant de vous offrir des services inovants dans ces suites bureautique.  
Cette extension vous permet d'envoyer du courrier électronique dans LibreOffice / OpenOffice, par un nouveau client smtp qui agit comme un serveur.

Etant un logiciel libre je vous encourage:
- A dupliquer son [code source](https://github.com/prrvchr/smtpServerOOo).
- A apporter des modifications, des corrections, des ameliorations.
- D'ouvrir un [dysfonctionnement](https://github.com/prrvchr/smtpServerOOo/issues/new) si nécessaire.
>>>>>>> smtpServerOOo/main

Bref, à participer au developpement de cette extension.
Car c'est ensemble que nous pouvons rendre le Logiciel Libre plus intelligent.

<<<<<<< HEAD
=======
## Prérequis:

smtpServerOOo utilise une base de données locale HsqlDB version 2.5.1.  
L'utilisation de HsqlDB nécessite l'installation et la configuration dans LibreOffice / OpenOffice d'un **JRE version 1.8 minimum** (c'est-à-dire: Java version 8)  
Je vous recommande [AdoptOpenJDK](https://adoptopenjdk.net/) comme source d'installation de Java.

Si vous utilisez **LibreOffice sous Linux**, alors vous êtes sujet au [dysfonctionnement 139538](https://bugs.documentfoundation.org/show_bug.cgi?id=139538).  
Pour contourner le problème, veuillez désinstaller les paquets:
- libreoffice-sdbc-hsqldb
- libhsqldb1.8.0-java

Si vous souhaitez quand même utiliser la fonctionnalité HsqlDB intégré fournie par LibreOffice, alors installez l'extension [HsqlDBembeddedOOo](https://prrvchr.github.io/HsqlDBembeddedOOo/README_fr).  
OpenOffice et LibreOffice sous Windows ne sont pas soumis à ce dysfonctionnement.

>>>>>>> smtpServerOOo/main
## Installation:

Il semble important que le fichier n'ait pas été renommé lors de son téléchargement.  
Si nécessaire, renommez-le avant de l'installer.

- Installer l'extension [OAuth2OOo.oxt](https://github.com/prrvchr/OAuth2OOo/raw/master/OAuth2OOo.oxt) version 0.0.5.

Vous devez d'abord installer cette extension, si elle n'est pas déjà installée.

<<<<<<< HEAD
- Installer l'extension [smtpMailerOOo.oxt](https://github.com/prrvchr/smtpMailerOOo/raw/master/smtpMailerOOo.oxt) version 0.0.1.
=======
- Installer l'extension [HsqlDBDriverOOo.oxt](https://github.com/prrvchr/HsqlDBDriverOOo/raw/master/HsqlDBDriverOOo.oxt) version 0.0.4.

Cette extension est nécessaire pour utiliser HsqlDB version 2.5.1 avec toutes ses fonctionnalités.

- Installer l'extension [smtpServerOOo.oxt](https://github.com/prrvchr/smtpServerOOo/raw/main/smtpServerOOo.oxt) version 0.0.1.
>>>>>>> smtpServerOOo/main

Redémarrez LibreOffice / OpenOffice après l'installation.

## Historique:

**Ce qui reste à faire:**

- Réécriture de mailmerge.py (pour être compatible avec: SSL, StartTLS et authentification OAuth2 ... c'est à dire: avec la technologie Mozilla IspBD).
- Ecrire un assistant à l'aide de la technologie Mozilla IspDB capable de trouver la configuration correcte fonctionnant avec mailmerge.py.
<<<<<<< HEAD
- Réécriture de Mailer Dialog.py en un assistant.
=======
- Ecriture d'un Service UNO, tournant en tâche de fond (Python Thread), permettant d'envoyer du courrier électronique.
>>>>>>> smtpServerOOo/main
