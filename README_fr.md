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

Bref, à participer au developpement de cette extension.
Car c'est ensemble que nous pouvons rendre le Logiciel Libre plus intelligent.

## Prérequis:

smtpServerOOo utilise une base de données locale HsqlDB version 2.5.1.  
L'utilisation de HsqlDB nécessite l'installation et la configuration dans LibreOffice / OpenOffice d'un **JRE version 1.8 minimum** (c'est-à-dire: Java version 8)  
Je vous recommande [AdoptOpenJDK](https://adoptopenjdk.net/) comme source d'installation de Java.

Si vous utilisez LibreOffice sous Linux, alors vous êtes sujet au [dysfonctionnement 139538](https://bugs.documentfoundation.org/show_bug.cgi?id=139538).  
Pour contourner le problème, veuillez désinstaller les paquets:
- libreoffice-sdbc-hsqldb
- libhsqldb1.8.0-java

Si vous souhaitez quand même utiliser la fonctionnalité HsqlDB intégré fournie par LibreOffice, alors installez l'extension [HsqlDBembeddedOOo](https://prrvchr.github.io/HsqlDBembeddedOOo/README_fr).  
OpenOffice et LibreOffice sous Windows ne sont pas soumis à ce dysfonctionnement.

## Installation:

Il semble important que le fichier n'ait pas été renommé lors de son téléchargement.  
Si nécessaire, renommez-le avant de l'installer.

- Installer l'extension [OAuth2OOo.oxt](https://github.com/prrvchr/OAuth2OOo/raw/master/OAuth2OOo.oxt) version 0.0.5.

Vous devez d'abord installer cette extension, si elle n'est pas déjà installée.

- Installer l'extension [smtpServerOOo.oxt](https://github.com/prrvchr/smtpServerOOo/raw/main/smtpServerOOo.oxt) version 0.0.1.

Redémarrez LibreOffice / OpenOffice après l'installation.

## Historique:

**Ce qui reste à faire:**

- Réécriture de mailmerge.py (pour être compatible avec: SSL, StartTLS et authentification OAuth2 ... c'est à dire: avec la technologie Mozilla IspBD).
- Ecrire un assistant à l'aide de la technologie Mozilla IspDB capable de trouver la configuration correcte fonctionnant avec mailmerge.py.
- Ecriture d'un Service UNO, tournant en tâche de fond (Python Thread), permettant d'envoyer du courrier électronique.
