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
# Documentation

**This [document][1] in english.**

**L'utilisation de ce logiciel vous soumet à nos [Conditions d'utilisation][2] et à notre [Politique de protection des données][3].**

# version [1.1.1][4]

## Introduction:

**eMailerOOo** fait partie d'une [Suite][5] d'extensions [LibreOffice][6] ~~et/ou [OpenOffice][7]~~ permettant de vous offrir des services inovants dans ces suites bureautique.  

Cette extension vous permet d'envoyer des documents dans LibreOffice sous forme de courriel, éventuellement par publipostage, à vos contacts téléphoniques.  
Elle fournit en plus une API utilisable en BASIC permettant d'envoyer des courriels et supportant les technologies les plus avancées (protocole OAuth2, Mozilla IspDB, HTTP au lieu de SMTP/IMAP, ...).  

Etant un logiciel libre je vous encourage:
- A dupliquer son [code source][8].
- A apporter des modifications, des corrections, des améliorations.
- D'ouvrir un [dysfonctionnement][9] si nécessaire.

Bref, à participer au developpement de cette extension.  
Car c'est ensemble que nous pouvons rendre le Logiciel Libre plus intelligent.

___
## Prérequis:

Afin de profiter des dernières versions des bibliothèques Python utilisées dans eMailerOOo, la version 2 de Python a été abandonnée au profit de **Python 3.8 minimum**.  
Cela signifie que **eMailerOOo ne supporte plus OpenOffice et LibreOffice 6.x sous Windows depuis sa version 1.0.0**.
Je ne peux que vous conseiller **de migrer vers LibreOffice 7.x**.

eMailerOOo utilise une base de données locale [HsqlDB][10] version 2.7.2.  
HsqlDB étant une base de données écrite en Java, son utilisation nécessite [l'installation et la configuration][11] dans LibreOffice / OpenOffice d'un **JRE version 11 ou ultérieure**.  
Je vous recommande [Adoptium][12] comme source d'installation de Java.

Si vous utilisez **LibreOffice Community sous Linux**, vous êtes sujet au [dysfonctionnement 139538][13]. Pour contourner le problème, veuillez **désinstaller les paquets** avec les commandes:
- `sudo apt remove libreoffice-sdbc-hsqldb` (pour désinstaller le paquet libreoffice-sdbc-hsqldb)
- `sudo apt remove libhsqldb1.8.0-java` (pour désinstaller le paquet libhsqldb1.8.0-java)

Si vous souhaitez quand même utiliser la fonctionnalité HsqlDB intégré fournie par LibreOffice, alors installez l'extension [HyperSQLOOo][14].  

___
## Installation:

Il semble important que le fichier n'ait pas été renommé lors de son téléchargement.  
Si nécessaire, renommez-le avant de l'installer.

- Installer l'extension ![OAuth2OOo logo][15] **[OAuth2OOo.oxt][16]** version 1.2.1. [![Download][17]][16]  
Vous devez d'abord installer cette extension, si elle n'est pas déjà installée.

- Installer l'extension ![jdbcDriverOOo logo][18] **[jdbcDriverOOo.oxt][19]** version 1.0.5. [![Download][20]][19]  
Cette extension est nécessaire pour utiliser HsqlDB version 2.7.2 avec toutes ses fonctionnalités.

- Si vous n'avez pas de source de données, vous pouvez installer l'une des extensions suivantes:

  - ![vCardOOo logo][21] **[vCardOOo.oxt][22]** version 1.0.3. [![Download][23]][22]  
  Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts présents sur une plateforme [**Nextcloud**][24] comme source de données pour les listes de diffusion et la fusion de documents.

  - ![gContactOOo logo][25] **[gContactOOo.oxt][26]** version 1.0.3. [![Download][27]][26]  
  Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts téléphoniques personnels (contact Android) comme source de données pour les listes de diffusion et la fusion de documents.

  - ![mContactOOo logo][28] **[mContactOOo.oxt][29]** version 1.0.3. [![Download][30]][29]  
  Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts Microsoft Outlook comme source de données pour les listes de diffusion et la fusion de documents.

- Installer l'extension ![eMailerOOo logo][31] **[eMailerOOo.oxt][32]** version  [![Version][33]][32]

Redémarrez LibreOffice / OpenOffice après l'installation.

___
## Utilisation:

### Introduction:

Pour pouvoir utiliser la fonctionnalité de publipostage de courriels en utilisant des listes de diffusion, il est nécessaire d'avoir une **source de données** avec des tables ayant les colonnes suivantes:
- Une ou plusieurs colonnes d'adresses électroniques. Ces colonnes sont choisi dans une liste et si ce choix n'est pas unique, alors la première colonne d'adresse courriel non nulle sera utilisée.
- Une ou plusieurs colonnes de clé primaire permettant d'identifier de manière unique les enregistrements, elle peut être une clé primaire composée. Les types supportés sont VARCHAR et/ou INTEGER, ou derivé. Ces colonnes doivent être declarée avec la contrainte NOT NULL.

De plus, cette **source de données** doit avoir au moins une **table principale**, comprenant tous les enregistrements pouvant être utilisés lors du publipostage du courriel.

Si vous ne disposez pas d'une telle **source de données** alors je vous invite à installer une des extensions suivantes :
- [vCardOOo][34]. Cette extension vous permettra d'utiliser vos contacts présents sur une plateforme [**Nextcloud**][24] comme source de données.
- [gContactOOo][35]. Cette extension vous permettra d'utiliser votre téléphone Android (vos contacts téléphoniques) comme source de données.
- [mContactOOo][36]. Cette extension vous permettra d'utiliser vos contacts Microsoft Outlook comme source de données.

Pour ces 3 extensions le nom de la **table principale** peut être trouvé (et même changé avant toute connexion) dans:  
**Outils -> Options -> Internet -> Nom de l'extension -> Nom de la table principale**

Ce mode d'utilisation est composé de 3 sections:
- [Publipostage de courriels avec des listes de diffusion][37].
- [Configuration de la connexion][38].
- [Courriels sortants][39].

### Publipostage de courriels avec des listes de diffusion:

#### Prérequis:

Pour pouvoir publiposter des courriels suivant une liste de diffusion, vous devez:
- Disposer d'une **source de données** comme décrit dans l'introduction précédente.
- Ouvrir un **nouveau document** Writer dans LibreOffice / OpenOffice.

Ce document Writer peut inclure des champs de fusion (insérables par la commande: **Insertion -> Champ -> Autres champs -> Base de données -> Champ de publipostage**), cela est même nécessaire si vous souhaitez pouvoir personnaliser le contenu du courriel et d'eventuel fichiers attachés.  
Ces champs de fusion doivent uniquement faire référence à la **table principale** de la **source de données**.

Si vous souhaitez utiliser un **document Writer déja existant**, vous devez vous assurer en plus que la **source de données** et la **table principale** sont bien rattachées au document dans : **Outils -> Source du carnet d'adresses...**.

Si ces recommandations ne sont pas suivies alors **la fusion de documents ne fonctionnera pas** et ceci silencieusement.

#### Démarrage de l'assistant de publipostage de courriels:

Dans un document LibreOffice / OpenOffice Writer aller à: **Outils -> Add-ons -> Envoi de courriels -> Publiposter un document**

![eMailerOOo Merger screenshot 1][40]

#### Sélection de la source de données:

Le chargement de la source de données de l'assistant **Publipostage de courriels** devrait apparaître :

![eMailerOOo Merger screenshot 2][41]

Les captures d'écran suivantes utilisent l'extension [gContactOOo][35] comme **source de données**. Si vous utilisez votre propre **source de données**, il est nécessaire d'adapter les paramètres par rapport à celle-ci. 

Dans la copie d'écran suivante, on peut voir que la **source de données** gContactOOo s'appelle: `Adresses` et que dans la liste des tables la table: `PUBLIC.Tous mes contacts` est sélectionnée.

![eMailerOOo Merger screenshot 3][42]

Si aucune liste de diffusion n'existe, vous devez en créer une, en saisissant son nom et en validant avec: `ENTRÉE` ou le bouton `Ajouter`.

Assurez-vous lors de la création de la liste de diffusion que la **table principale** est toujours bien sélectionnée dans la liste des tables.  
Si cette recommandation n'est pas suivie alors **la fusion de documents ne fonctionnera pas** et ceci silencieusement.

![eMailerOOo Merger screenshot 4][43]

Maintenant que votre nouvelle liste de diffusion est disponible dans la liste, vous devez la sélectionner.

Et ajouter les colonnes suivantes:
- Colonne de clef primaire: `Uri`
- Colonnes d'adresses électronique: `HomeEmail`, `WorkEmail` et `OtherEmail`

Si plusieurs colonnes d'adresses courriel sont sélectionnées, alors l'ordre devient pertinent puisque le courriel sera envoyé à la première adresse disponible.  
De plus, à l'étape Sélection des destinataires de l'assistant, dans l'onglet [Destinataires disponibles][44], seuls les enregistrements avec au moins une colonne d'adresse courriel saisie seront répertoriés.  
Assurez-vous donc d'avoir un carnet d'adresses avec au moins un des champs d'adresse e-mail (Home, Work ou Other) renseigné.

![eMailerOOo Merger screenshot 5][45]

Ce paramètrage ne doit être effectué que pour les nouvelles listes de diffusion.  
Vous pouvez maintenant passer à l'étape suivante.

#### Sélection des destinataires:

##### Destinataires disponibles:

Les destinataires sont sélectionnés à l'aide de 2 boutons `Tout ajouter` et `Ajouter` permettant respectivement:
- Soit d'ajouter le groupe de destinataires sélectionnés dans la liste `Carnet d'adresses`. Ceci permet lors d'un publipostage, que les modifications du contenu du groupe soient prises en compte. Une liste de diffusion n'accepte qu'un seul groupe.
- Soit d'ajouter la sélection, qui peut être multiple à l'aide de la touche `CTRL`. Cette sélection est immuable quelle que soit la modification des groupes du carnet d'adresses.

![eMailerOOo Merger screenshot 6][46]

Example de la sélection multiple:

![eMailerOOo Merger screenshot 7][47]

##### Destinataires sélectionnés:

Les destinataires sont désélectionnés à l'aide de 2 boutons `Tout retirer` et `Retirer` permettant respectivement:
- Soit de retirer le groupe qui a été affecté à cette liste de diffusion. Ceci est nécessaire afin de pouvoir modifier à nouveau le contenu de cette liste de diffusion.
- Soit de retirer la sélection, qui peut être multiple à l'aide de la touche `CTRL`.

![eMailerOOo Merger screenshot 8][48]

Si vous avez sélectionné au moins 1 destinataire, vous pouvez passer à l'étape suivante.

#### Sélection des options d'envoi:

Si cela n'est pas déjà fait, vous devez créer un nouvel expéditeur à l'aide du bouton `Ajouter`.

![eMailerOOo Merger screenshot 9][49]

La création du nouvel expéditeur est décrite dans la section [Configuration de la connexion][38].

Le courriel doit avoir un sujet. Il peut être enregistré dans le document Writer.  
Vous pouvez insérer des champs de fusion dans l'objet du courriel. Un champ de fusion est composé d'une accolade ouvrante, du nom de la colonne référencée (sensible à la casse) et d'une accolade fermante (ie: `{NomColonne}`).

![eMailerOOo Merger screenshot 10][50]

Le courriel peut éventuellement contenir des fichiers joints. Ils peuvent être enregistrés dans le document Writer.  
La capture d'écran suivante montre 1 fichier joint qui sera fusionné sur la source de données puis converti au format PDF avant d'être joint au courriel.

![eMailerOOo Merger screenshot 11][51]

Assurez-vous de toujours quitter l'assistant avec le bouton `Terminer` pour confirmer la soumission des travaux d'envoi.  
Pour envoyer les travaux d'envoi, veuillez suivre la section [Courriels sortants][39].

### Configuration de la connexion:

#### Démarrage de l'assistant de connexion:

Dans LibreOffice / OpenOffice aller à: **Outils -> Add-ons -> Envoi de courriels -> Configurer la connexion**

![eMailerOOo Ispdb screenshot 1][52]

#### Sélection du compte:

![eMailerOOo Ispdb screenshot 2][53]

#### Trouver la configuration:

![eMailerOOo Ispdb screenshot 3][54]

#### Configuration SMTP:

![eMailerOOo Ispdb screenshot 4][55]

#### Configuration IMAP:

![eMailerOOo Ispdb screenshot 5][56]

#### Tester la connexion:

![eMailerOOo Ispdb screenshot 6][57]

Assurez-vous de toujours quitter l'assistant avec le bouton `Terminer` afin d'enregistrer les paramètres de connexion.

### Courriels sortants:

#### Démarrage du spouleur de courriels:

Dans LibreOffice / OpenOffice aller à: **Outils -> Add-ons -> Envoi de courriels -> Courriels sortants**

![eMailerOOo Spooler screenshot 1][58]

#### Liste des courriels sortants:

Chaque travaux d'envoi possède 3 états différents:
- État **0**: le courriel est prêt à être envoyé.
- État **1**: le courriel a été envoyé avec succès.
- État **2**: Une erreur est survenue lors de l'envoi du courriel. Vous pouvez consulter le message d'erreur dans le [Journal d'activité du spouleur][59].

![eMailerOOo Spooler screenshot 2][60]

Le spouleur de courriels est arrêté par défaut. **Il doit être démarré avec le bouton `Démarrer / Arrêter` pour que les courriels en attente soient envoyés**.

#### Journal d'activité du spouleur:

Lorsque le spouleur de courriel est démarré, son activité peut être visualisée dans le journal d'activité.

![eMailerOOo Spooler screenshot 3][61]

___
## A été testé avec:

* LibreOffice 7.3.7.2 - Lubuntu 22.04 - Python version 3.10.12 - OpenJDK-11-JRE (amd64)

* LibreOffice 7.5.4.2(x86) - Windows 10 - Python version 3.8.16 - Adoptium JDK Hotspot 11.0.19 (under Lubuntu 22.04 / VirtualBox 6.1.38)

* LibreOffice 7.4.3.2(x64) - Windows 10(x64) - Python version 3.8.15  - Adoptium JDK Hotspot 11.0.17 (x64) (under Lubuntu 22.04 / VirtualBox 6.1.38)

* **Ne fonctionne pas avec OpenOffice sous Windows** voir [dysfonctionnement 128569][62]. N'ayant aucune solution, je vous encourrage d'installer **LibreOffice**.

Je vous encourage en cas de problème :confused:  
de créer un [dysfonctionnement][9]  
J'essaierai de le résoudre :smile:

___
## Historique:

### Ce qui a été fait pour la version 0.0.1:

- Ecriture de [IspDB][63] ou l'assistant de Configuration de connexion aux serveurs SMTP permettant:
    - De trouver les paramètres de connexion à un serveur SMTP à partir d'une adresse courriel. D'ailleur je remercie particulierement Mozilla, pour [Thunderbird autoconfiguration database][64] ou IspDB, qui à rendu ce défi possible...
    - D'afficher l'activité du service UNO `com.sun.star.mail.MailServiceProvider` lors de la connexion au serveur SMTP et l'envoi d'un courriel. 

- Ecriture du [Spouleur][65] de courriels permettant:
    - D'afficher les travaux d'envoi de courriel avec leurs états respectifs.
    - D'afficher l'activité du service UNO `com.sun.star.mail.SpoolerService` lors de l'envoi de courriels.
    - De démarrer et arrêter le service spouleur.

- Ecriture du [Merger][66] ou l'assistant de publipostage de courriels permettant:
    - De créer des listes de diffusions.
    - De fusionner et convertir au format HTML le document courant pour en faire le message du courriel.
    - De fusionner et/ou convertir au format PDF d'éventuel fichiers joints au courriel.

- Ecriture du [Mailer][67] de document permettant:
    - De convertir au format HTML le document pour en faire le message du courriel.
    - De convertir au format PDF d'éventuel fichiers joints au courriel.

- Ecriture d'un [Grid][68] piloté par un `com.sun.star.sdb.RowSet` permettant:
    - D'être paramètrable sur les colonnes à afficher.
    - D'être paramètrable sur l'ordre de tri à afficher.
    - De sauvegarder les paramètres d'affichage.

### Ce qui a été fait pour la version 0.0.2:

- Réécriture de [IspDB][63] ou Assistant de configuration de connexion aux serveurs de messagerie afin d'intégrer la configuration de la connexion IMAP.
  - Utilisation de [IMAPClient][69] version 2.2.0: une bibliothèque cliente IMAP complète, Pythonic et facile à utiliser.
  - Extension des fichiers IDL [com.sun.star.mail.*][70]:
    - [XMailMessage2.idl][71] prend désormais en charge la hiérarchisation des courriels (thread).
    - La nouvelle interface [XImapService][72] permet d'accéder à une partie de la bibliothèque IMAPClient.

- Réécriture du [Spouleur][73] afin d'intégrer des fonctionnalités IMAP comme la création d'un fil récapitulant le publipostage et regroupant tous les courriels envoyés.

- Soumission de l'extension eMailerOOo à Google et obtention de l'autorisation d'utiliser son API GMail afin d'envoyer des courriels avec un compte Google.

### Ce qui a été fait pour la version 0.0.3:

- Réécriture du [Grid][68] afin de permettre:
  - Le tri sur une colonne avec l'intégration du service UNO [SortableGridDataModel][74].
  - La génération des filtres des enregistrements nécessaires au service [Spouleur][65].
  - Le partage du module python avec l'extension [jdbcDriverOOo][75].

- Réécriture du [Merger][66] afin de permettre:
  - La gestion du nom du Schema dans de nom des tables afin d'être compatible avec la version 0.0.4 de [jdbcDriverOOo][76].
  - La création de liste de diffusion sur un groupe du carnet d'adresse et permettant de suivre la modification de son contenu.
  - L'utilisation de clé primaire, qui peuvent être composite, supportant les [DataType][77] `VARCHAR` et `INTEGER` ou derivé.
  - Un aperçu du document avec des champs de fusion remplis plus rapidement grâce au [Grid][68].

- Réécriture du [Spouleur][65] afin de permettre:
  - L'utilisation des nouveaux filtres supportant les clés primaires composite fourni par le [Merger][66].
  - L'utilisation du nouveau [Grid][68] permettant le tri sur une colonne.

- Encore plein d'autres choses...

### Ce qui a été fait pour la version 1.0.0:

- L'extension **smtpMailerOOo** a été renomé en **eMailerOOo**.

### Ce qui a été fait pour la version 1.0.1:

- L'absence ou l'obsolescence des extensions **OAuth2OOo** et/ou **jdbcDriverOOo** nécessaires au bon fonctionnement de **eMailerOOo** affiche désormais un message d'erreur. Ceci afin d'éviter qu'un dysfonctionnement tel que le [dysfonctionnement #3][78] ne se reproduise...

- La base de données HsqlDB sous-jacente peut être ouverte dans Base avec: **Outils -> Options -> Internet -> eMailerOOo -> Base de données**.

- Le menu **Outils -> Add-ons** s'affiche désormais correctement en fonction du contexte.

- Encore plein d'autres choses...

### Ce qui a été fait pour la version 1.0.2:

- Si aucune configuration n'est trouvée dans l'assistant de configuration de la connexion (IspDB Wizard) alors il est possible de configurer la connexion manuellement. Voir [dysfonctionnement #5][79].

### Ce qui a été fait pour la version 1.1.0:

- Dans l'assistant de configuration de la connexion (IspDB Wizard) il est maintenant possible de désactiver la configuration IMAP.  
  En conséquence, cela n'envoie plus de fil de discussion (message IMAP) lors de la fusion d'un mailing.  
  Dans ce même assistant, il est désormais possible de saisir une adresse courriel de réponse.

- Dans l'assistant de fusion d'email, il est désormais possible d'insérer des champs de fusion dans l'objet du courriel. Voir [dysfonctionnement #6][80].  
  Dans le sujet d'un courriel, un champ de fusion est composé d'une accolade ouvrante, du nom de la colonne référencée (sensible à la casse) et d'une accolade fermante (ie: `{NomDeLaColonne}`).  
  Lors de la saisie du sujet du courriel, une erreur de syntaxe dans un champ de fusion sera signalée et empêchera la soumission du mailing.

- Il est désormais possible dans le Spooler de visualiser les courriels au format eml.

- Un service [com.sun.star.mail.MailUser][81] permet désormais d'accéder à une configuration de connexion (SMTP et/ou IMAP) depuis une adresse courriel qui suite la rfc822.  
  Un autre service [com.sun.star.datatransfer.TransferableFactory][82] permet, comme son nom l'indique, la création de [Transferable][83] à partir d'un texte (string), d'une séquence binaire, d'une Url (file://...) ou un flux de données (InputStream).  
  Ces deux nouveaux services simplifient grandement l'API mail de LibreOffice et permettent d'envoyer des courriels depuis Basic. Voir le [dysfonctionnement #4][84].  
  Vous trouverez une macro Basic vous permettant d'envoyer des emails dans : **Outils -> Macros -> Editer les Macros... -> eMailerOOo -> SendEmail**.

### Ce qui a été fait pour la version 1.1.1:

- Prise en charge de la version 1.2.0 de l'extension **OAuth2OOo**. Les versions précédentes ne fonctionneront pas avec l'extension **OAuth2OOo** 1.2.0 ou ultérieure.

### Que reste-t-il à faire pour la version 1.1.1:

- Ajouter de nouvelles langues pour l’internationalisation...

- Tout ce qui est bienvenu...

[1]: <https://prrvchr.github.io/eMailerOOo/>
[2]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/TermsOfUse_fr>
[3]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/PrivacyPolicy_fr>
[4]: <https://prrvchr.github.io/eMailerOOo/README_fr#ce-qui-a-%C3%A9t%C3%A9-fait-pour-la-version-110>
[5]: <https://prrvchr.github.io/>
[6]: <https://www.libreoffice.org/download/download-libreoffice/>
[7]: <https://www.openoffice.org/download/index.html>
[8]: <https://github.com/prrvchr/eMailerOOo>
[9]: <https://github.com/prrvchr/eMailerOOo/issues/new>
[10]: <http://hsqldb.org/>
[11]: <https://wiki.documentfoundation.org/Documentation/HowTo/Install_the_correct_JRE_-_LibreOffice_on_Windows_10/fr>
[12]: <https://adoptium.net/releases.html?variant=openjdk11>
[13]: <https://bugs.documentfoundation.org/show_bug.cgi?id=139538>
[14]: <https://prrvchr.github.io/HyperSQLOOo/>
[15]: <https://prrvchr.github.io/OAuth2OOo/img/OAuth2OOo.svg>
[16]: <https://github.com/prrvchr/OAuth2OOo/releases/latest/download/OAuth2OOo.oxt>
[17]: <https://img.shields.io/github/downloads/prrvchr/OAuth2OOo/latest/total#right>
[18]: <https://prrvchr.github.io/jdbcDriverOOo/img/jdbcDriverOOo.svg>
[19]: <https://github.com/prrvchr/jdbcDriverOOo/releases/latest/download/jdbcDriverOOo.oxt>
[20]: <https://img.shields.io/github/downloads/prrvchr/jdbcDriverOOo/latest/total#right>
[21]: <https://prrvchr.github.io/vCardOOo/img/vCardOOo.svg>
[22]: <https://github.com/prrvchr/vCardOOo/releases/latest/download/vCardOOo.oxt>
[23]: <https://img.shields.io/github/downloads/prrvchr/vCardOOo/latest/total#right>
[24]: <https://fr.wikipedia.org/wiki/Nextcloud>
[25]: <https://prrvchr.github.io/gContactOOo/img/gContactOOo.svg>
[26]: <https://github.com/prrvchr/gContactOOo/releases/latest/download/gContactOOo.oxt>
[27]: <https://img.shields.io/github/downloads/prrvchr/gContactOOo/latest/total#right>
[28]: <https://prrvchr.github.io/mContactOOo/img/mContactOOo.svg>
[29]: <https://github.com/prrvchr/mContactOOo/releases/latest/download/mContactOOo.oxt>
[30]: <https://img.shields.io/github/downloads/prrvchr/mContactOOo/latest/total#right>
[31]: <https://prrvchr.github.io/eMailerOOo/img/eMailerOOo.svg>
[32]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/eMailerOOo.oxt>
[33]: <https://img.shields.io/github/downloads/prrvchr/eMailerOOo/latest/total#right>
[34]: <https://prrvchr.github.io/vCardOOo/README_fr>
[35]: <https://prrvchr.github.io/gContactOOo/README_fr>
[36]: <https://prrvchr.github.io/mContactOOo/README_fr>
[37]: <https://prrvchr.github.io/eMailerOOo/README_fr#publipostage-de-courriels-avec-des-listes-de-diffusion>
[38]: <https://prrvchr.github.io/eMailerOOo/README_fr#configuration-de-la-connexion>
[39]: <https://prrvchr.github.io/eMailerOOo/README_fr#courriels-sortants>
[40]: <img/eMailerOOo-Merger1.png>
[41]: <img/eMailerOOo-Merger2.png>
[42]: <img/eMailerOOo-Merger3.png>
[43]: <img/eMailerOOo-Merger4.png>
[44]: <https://prrvchr.github.io/eMailerOOo/README_fr#destinataires-disponibles>
[45]: <img/eMailerOOo-Merger5.png>
[46]: <img/eMailerOOo-Merger6.png>
[47]: <img/eMailerOOo-Merger7.png>
[48]: <img/eMailerOOo-Merger8.png>
[49]: <img/eMailerOOo-Merger9.png>
[50]: <img/eMailerOOo-Merger10.png>
[51]: <img/eMailerOOo-Merger11.png>
[52]: <img/eMailerOOo-Ispdb1.png>
[53]: <img/eMailerOOo-Ispdb2.png>
[54]: <img/eMailerOOo-Ispdb3.png>
[55]: <img/eMailerOOo-Ispdb4.png>
[56]: <img/eMailerOOo-Ispdb5.png>
[57]: <img/eMailerOOo-Ispdb6.png>
[58]: <img/eMailerOOo-Spooler1.png>
[59]: <https://prrvchr.github.io/eMailerOOo/README_fr#journal-dactivité-du-spouleur>
[60]: <img/eMailerOOo-Spooler2.png>
[61]: <img/eMailerOOo-Spooler3.png>
[62]: <https://bz.apache.org/ooo/show_bug.cgi?id=128569>
[63]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/ispdb>
[64]: <https://wiki.mozilla.org/Thunderbird:Autoconfiguration>
[65]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler>
[66]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/merger>
[67]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/mailer>
[68]: <https://github.com/prrvchr/eMailerOOo/tree/master/uno/lib/uno/grid>
[69]: <https://github.com/mjs/imapclient#readme>
[70]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/idl/com/sun/star/mail>
[71]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailMessage2.idl>
[72]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XImapService.idl>
[73]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/spooler.py>
[74]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/grid/SortableGridDataModel.html>
[75]: <https://github.com/prrvchr/jdbcDriverOOo/tree/master/source/jdbcDriverOOo/service/pythonpath/jdbcdriver/grid>
[76]: <https://prrvchr.github.io/jdbcDriverOOo/>
[77]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/sdbc/DataType.html>
[78]: <https://github.com/prrvchr/eMailerOOo/issues/3>
[79]: <https://github.com/prrvchr/eMailerOOo/issues/5>
[80]: <https://github.com/prrvchr/eMailerOOo/issues/6>
[81]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailUser.idl>
[82]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/datatransfer/XTransferableFactory.idl>
[83]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/datatransfer/XTransferable.html>
[84]: <https://github.com/prrvchr/eMailerOOo/issues/4>
