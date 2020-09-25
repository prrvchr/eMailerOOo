**Cette extension est en cours de développement et n'est pas encore disponible ... Merci de votre patience.**

# version [v0.0.1](https://prrvchr.github.io/smtpMailerOOo/README_fr#historique)

[**This document in English**](https://prrvchr.github.io/smtpMailerOOo)

**L'utilisation de ce logiciel vous soumet à nos** [**Conditions d'utilisation**](https://prrvchr.github.io/smtpMailerOOo/smtpMailerOOo/registration/TermsOfUse_fr) **et à notre** [**Politique de protection des données**](https://prrvchr.github.io/smtpMailerOOo/smtpMailerOOo/registration/PrivacyPolicy_fr)

## Introduction:

**smtpMailerOOo** est une extension [LibreOffice](https://fr.libreoffice.org/download/telecharger-libreoffice/) et/ou [OpenOffice](https://www.openoffice.org/fr/Telecharger/) permettant de vous offrir des services inovants dans ces suites bureautique.  
Cette extension vous permet d'envoyer du courrier électronique dans LibreOffice / OpenOffice, éventuellement par publipostage, à vos contacts téléphoniques.

Etant un logiciel libre je vous encourage:
- A dupliquer son [code source](https://github.com/prrvchr/smtpMailerOOo).
- A apporter des modifications, des corrections, des ameliorations.

Bref, à participer au developpement de cette extension.
Car c'est ensemble que nous pouvons rendre le Logiciel Libre plus intelligent.

## Installation:

Il semble important que le fichier n'ait pas été renommé lors de son téléchargement.  
Si nécessaire, renommez-le avant de l'installer.

- Installer l'extension [OAuth2OOo.oxt](https://github.com/prrvchr/OAuth2OOo/raw/master/OAuth2OOo.oxt) version 0.0.5.

Vous devez d'abord installer cette extension, si elle n'est pas déjà installée.

- Installer l'extension [smtpMailerOOo.oxt](https://github.com/prrvchr/smtpMailerOOo/raw/master/smtpMailerOOo.oxt) version 0.0.1.

Redémarrez LibreOffice / OpenOffice après l'installation.

## Historique:

**Ce qui reste à faire:**

- Réécriture de mailmerge.py (pour être compatible avec: SSL, StartTLS et authentification OAuth2 ... c'est à dire: avec la technologie Mozilla IspBD).
- Ecrire un assistant à l'aide de la technologie Mozilla IspDB capable de trouver la configuration correcte fonctionnant avec mailmerge.py.
- Réécriture de Mailer Dialog.py en un assistant.
