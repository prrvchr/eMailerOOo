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
- D'ouvrir une [issue](https://github.com/prrvchr/smtpServerOOo/issues/new) si nécessaire.

Bref, à participer au developpement de cette extension.
Car c'est ensemble que nous pouvons rendre le Logiciel Libre plus intelligent.

## Prérequis:

smtpServerOOo utilise une base de données locale HsqlDB version 2.5.1.  
L'utilisation de HsqlDB nécessite l'installation et la configuration dans
LibreOffice / OpenOffice d'un **JRE version 1.8 minimum** (c'est-à-dire: Java version 8)

Parfois, il peut être nécessaire pour les utilisateurs de LibreOffice de ne pas avoir de pilote HsqlDB installé avec LibreOffice  
(vérifiez vos applications installées sous Windows ou votre gestionnaire de paquets sous Linux)  
~~Il semble que les versions 6.4.x et 7.x de LibreOffice aient résolu ce problème et sont capables de fonctionner simultanément avec différentes versions de pilote de HsqlDB.~~  
Après de nombreux tests, il semble que LibreOffice (6.4.x et 7.x) ne puisse pas charger un pilote HsqlDB fourni (hsqldb.jar v2.5.1), si le pilote HsqlDB intégré est installé (et même la solution est parfois de renommer le fichier hsqldb.jar dans /usr/share/java, la désinstallation du paquet libreoffice-sdbc-hsqldb ne semble pas suffisante...)  
Pour surmonter cette limitation et si vous souhaitez utiliser HsqlDB intégré, supprimez le pilote HsqlDB intégré (hsqldb.jar v1.8.0) et installez cette extension: [HsqlDBembeddedOOo](https://prrvchr.github.io/HsqlDBembeddedOOo/) pour remplacer le pilote HsqlDB intégré disfonctionnant de LibreOffice.  
OpenOffice ne semble pas avoir besoin de cette solution de contournement.

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
