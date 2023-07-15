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

**This [document][2] in english.**

**L'utilisation de ce logiciel vous soumet à nos [Conditions d'utilisation][3] et à notre [Politique de protection des données][4].**

# version [1.0.0][5]

## Introduction:

**eMailerOOo** fait partie d'une [Suite][6] d'extensions [LibreOffice][7] ~~et/ou [OpenOffice][8]~~ permettant de vous offrir des services inovants dans ces suites bureautique.  
Cette extension vous permet d'envoyer des documents dans LibreOffice / OpenOffice sous forme de courriel, éventuellement par publipostage, à vos contacts téléphoniques.

Etant un logiciel libre je vous encourage:
- A dupliquer son [code source][9].
- A apporter des modifications, des corrections, des améliorations.
- D'ouvrir un [dysfonctionnement][10] si nécessaire.

Bref, à participer au developpement de cette extension.  
Car c'est ensemble que nous pouvons rendre le Logiciel Libre plus intelligent.

## Prérequis:

Si vous utilisez **OpenOffice sous Windows** quelle que soit la version, vous êtes sujet au [dysfonctionnement 128569][11]. Je n'ai pas trouvé de solution de contournement, pour l'instant je ne peux que vous conseiller d'installer **LibreOffice**...

eMailerOOo utilise une base de données locale [HsqlDB][12] version 2.7.2.  
HsqlDB étant une base de données écrite en Java, son utilisation nécessite [l'installation et la configuration][13] dans LibreOffice / OpenOffice d'un **JRE version 11 ou ultérieure**.  
Je vous recommande [Adoptium][14] comme source d'installation de Java.

Si vous utilisez **LibreOffice sous Linux**, vous devez vous assurez de deux choses:
  - Vous êtes sujet au [dysfonctionnement 139538][15]. Pour contourner le problème, veuillez **désinstaller les paquets** avec les commandes:
    - `sudo apt remove libreoffice-sdbc-hsqldb` (pour désinstaller le paquet libreoffice-sdbc-hsqldb)
    - `sudo apt remove libhsqldb1.8.0-java` (pour désinstaller le paquet libhsqldb1.8.0-java)

Si vous souhaitez quand même utiliser la fonctionnalité HsqlDB intégré fournie par LibreOffice, alors installez l'extension [HsqlDBembeddedOOo][16].  

  - Si le paquet python3-cffi-backend est installé alors vous devez **installer le paquet python3-cffi** avec les commandes:
    - `dpkg -s python3-cffi-backend` (pour savoir si le paquet python3-cffi-backend est installé)
    - `sudo apt install python3-cffi` (pour installer le paquet python3-cffi si nécessaire)

OpenOffice sous Linux et LibreOffice sous Windows ne sont pas soumis à ces dysfonctionnements.

## Installation:

Il semble important que le fichier n'ait pas été renommé lors de son téléchargement.  
Si nécessaire, renommez-le avant de l'installer.

- Installer l'extension ![OAuth2OOo logo][17] **[OAuth2OOo.oxt][18]** version 1.1.0.  
Vous devez d'abord installer cette extension, si elle n'est pas déjà installée.

- Installer l'extension ![jdbcDriverOOo logo][19] **[jdbcDriverOOo.oxt][20]** version 1.0.0.  
Cette extension est nécessaire pour utiliser HsqlDB version 2.7.2 avec toutes ses fonctionnalités.

- Si vous n'avez pas de source de données, vous pouvez installer l'une des extensions suivantes:

  - ![vCardOOo logo][21] **[vCardOOo.oxt][22]** version 1.0.0.  
  Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts présents sur une plateforme [**Nextcloud**][23] comme source de données pour les listes de diffusion et la fusion de documents.

  - ![gContactOOo logo][24] **[gContactOOo.oxt][25]** version 1.0.0.  
  Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts téléphoniques personnels (contact Android) comme source de données pour les listes de diffusion et la fusion de documents.

  - ![mContactOOo logo][26] **[mContactOOo.oxt][27]** version 1.0.0.  
  Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts Microsoft Outlook comme source de données pour les listes de diffusion et la fusion de documents.

- Installer l'extension ![eMailerOOo logo][1] **[eMailerOOo.oxt][28]** version 1.0.0.  

Redémarrez LibreOffice / OpenOffice après l'installation.

## Utilisation:

### Introduction:

Pour pouvoir utiliser la fonctionnalité de publipostage de courriels en utilisant des listes de diffusion, il est nécessaire d'avoir une **source de données** avec des tables ayant les colonnes suivantes:
- Une ou plusieurs colonnes d'adresses électroniques. Ces colonnes sont choisi dans une liste et si ce choix n'est pas unique, alors la première colonne d'adresse courriel non nulle sera utilisée.
- Une ou plusieurs colonnes de clé primaire permettant d'identifier de manière unique les enregistrements, elle peut être une clé primaire composée. Les types supportés sont VARCHAR et/ou INTEGER, ou derivé. Ces colonnes doivent être declarée avec la contrainte NOT NULL.

De plus, cette **source de données** doit avoir au moins une **table principale**, comprenant tous les enregistrements pouvant être utilisés lors du publipostage du courriel.

Si vous ne disposez pas d'une telle **source de données** alors je vous invite à installer une des extensions suivantes :
- [vCardOOo][22]. Cette extension vous permettra d'utiliser vos contacts présents sur une plateforme [**Nextcloud**][23] comme source de données.
- [gContactOOo][25]. Cette extension vous permettra d'utiliser votre téléphone Android (vos contacts téléphoniques) comme source de données.
- [mContactOOo][27]. Cette extension vous permettra d'utiliser vos contacts Microsoft Outlook comme source de données.

Pour ces 3 extensions le nom de la **table principale** peut être trouvé (et même changé avant toute connexion) dans:  
**Outils -> Options -> Internet -> Nom de l'extension -> Nom de la table principale**

Ce mode d'utilisation est composé de 3 sections:
- [Publipostage de courriels avec des listes de diffusion][29].
- [Configuration de la connexion][30].
- [Courriels sortants][31].

### Publipostage de courriels avec des listes de diffusion:

#### Prérequis:

Pour pouvoir publiposter des courriels suivant une liste de diffusion, vous devez:
- Disposer d'une **source de données** comme décrit dans l'introduction précédente.
- Ouvrir un **nouveau document** Writer dans LibreOffice / OpenOffice.

Ce document Writer peut inclure des champs de fusion (insérables par la commande: **Insertion -> Champ -> Autres champs -> Base de données -> Champ de publipostage**), cela est même nécessaire si vous souhaitez pouvoir personnaliser le contenu du courriel et d'eventuel fichiers attachés.  
Ces champs de fusion doivent uniquement faire référence à la **table principale** de la **source de données**.

Si vous souhaitez utiliser un **document Writer déja existant**, vous devez vous assurer en plus que la **source de données** et la **table principale** sont bien rattachées au document dans : **Outils -> Source du carnet d'adresses...**.

Si ces recommandations ne sont pas suivies alors **la fusion de documents de fonctionnera pas** et ceci silencieusement.

#### Démarrage de l'assistant de publipostage de courriels:

Dans un document LibreOffice / OpenOffice Writer aller à: **Outils -> Add-ons -> Envoi de courriels -> Publiposter un document**

![eMailerOOo Merger screenshot 1][32]

#### Sélection de la source de données:

Le chargement de la source de données de l'assistant **Publipostage de courriels** devrait apparaître :

![eMailerOOo Merger screenshot 2][33]

Les captures d'écran suivantes utilisent l'extension [gContactOOo][25] comme **source de données**. Si vous utilisez votre propre **source de données**, il est nécessaire d'adapter les paramètres par rapport à celle-ci. 

Dans la copie d'écran suivante, on peut voir que la **source de données** gContactOOo s'appelle: `Adresses` et que dans la liste des tables la table: `PUBLIC.Tous mes contacts` est sélectionnée.

![eMailerOOo Merger screenshot 3][34]

Si aucune liste de diffusion n'existe, vous devez en créer une, en saisissant son nom et en validant avec: `ENTRÉE` ou le bouton `Ajouter`.

Assurez-vous lors de la création de la liste de diffusion que la **table principale** est toujours bien sélectionnée dans la liste des tables.  
Si cette recommandation n'est pas suivie alors **la fusion de documents de fonctionnera pas** et ceci silencieusement.

![eMailerOOo Merger screenshot 4][35]

Maintenant que votre nouvelle liste de diffusion est disponible dans la liste, vous devez la sélectionner.

Et ajouter les colonnes suivantes:
- Colonne de clef primaire: `Uri`
- Colonnes d'adresses électronique: `HomeEmail`, `WorkEmail` et `OtherEmail`

Si plusieurs colonnes d'adresses courriel sont sélectionnées, alors l'ordre devient pertinent puisque le courriel sera envoyé à la première adresse disponible.  
De plus, à l'étape Sélection des destinataires de l'assistant, dans l'onglet [Destinataires disponibles][36], seuls les enregistrements avec au moins une colonne d'adresse courriel saisie seront répertoriés.  
Assurez-vous donc d'avoir un carnet d'adresses avec au moins un des champs d'adresse e-mail (Home, Work ou Other) renseigné.

![eMailerOOo Merger screenshot 5][37]

Ce paramètrage ne doit être effectué que pour les nouvelles listes de diffusion.  
Vous pouvez maintenant passer à l'étape suivante.

#### Sélection des destinataires:

##### Destinataires disponibles:

Les destinataires sont sélectionnés à l'aide de 2 boutons `Tout ajouter` et `Ajouter` permettant respectivement:
- Soit d'ajouter le groupe de destinataires sélectionnés dans la liste `Carnet d'adresses`. Ceci permet lors d'un publipostage, que les modifications du contenu du groupe soient prises en compte. Une liste de diffusion n'accepte qu'un seul groupe.
- Soit d'ajouter la sélection, qui peut être multiple à l'aide de la touche `CTRL`. Cette sélection est immuable quelle que soit la modification des groupes du carnet d'adresses.

![eMailerOOo Merger screenshot 6][38]

Example de la sélection multiple:

![eMailerOOo Merger screenshot 7][39]

##### Destinataires sélectionnés:

Les destinataires sont désélectionnés à l'aide de 2 boutons `Tout retirer` et `Retirer` permettant respectivement:
- Soit de retirer le groupe qui a été affecté à cette liste de diffusion. Ceci est nécessaire afin de pouvoir modifier à nouveau le contenu de cette liste de diffusion.
- Soit de retirer la sélection, qui peut être multiple à l'aide de la touche `CTRL`.

![eMailerOOo Merger screenshot 8][40]

Si vous avez sélectionné au moins 1 destinataire, vous pouvez passer à l'étape suivante.

#### Sélection des options d'envoi:

Si cela n'est pas déjà fait, vous devez créer un nouvel expéditeur à l'aide du bouton `Ajouter`.

![eMailerOOo Merger screenshot 9][41]

La création du nouvel expéditeur est décrite dans la section [Configuration de la connexion][42].

Le courriel doit avoir un sujet. Il peut être enregistré dans le document Writer.

![eMailerOOo Merger screenshot 10][43]

Le courriel peut éventuellement contenir des fichiers joints. Ils peuvent être enregistrés dans le document Writer.  
La capture d'écran suivante montre 1 fichier joint qui sera fusionné sur la source de données puis converti au format PDF avant d'être joint au courriel.

![eMailerOOo Merger screenshot 11][44]

Assurez-vous de toujours quitter l'assistant avec le bouton `Terminer` pour confirmer la soumission des travaux d'envoi.  
Pour envoyer les travaux d'envoi, veuillez suivre la section [Courriels sortants][45].

### Configuration de la connexion:

#### Démarrage de l'assistant de connexion:

Dans LibreOffice / OpenOffice aller à: **Outils -> Add-ons -> Envoi de courriels -> Configurer la connexion**

![eMailerOOo Ispdb screenshot 1][46]

#### Sélection du compte:

![eMailerOOo Ispdb screenshot 2][47]

#### Trouver la configuration:

![eMailerOOo Ispdb screenshot 3][48]

#### Configuration SMTP:

![eMailerOOo Ispdb screenshot 4][49]

#### Configuration IMAP:

![eMailerOOo Ispdb screenshot 5][50]

#### Tester la connexion:

![eMailerOOo Ispdb screenshot 6][51]

Assurez-vous de toujours quitter l'assistant avec le bouton `Terminer` afin d'enregistrer les paramètres de connexion.

### Courriels sortants:

#### Démarrage du spouleur de courriels:

Dans LibreOffice / OpenOffice aller à: **Outils -> Add-ons -> Envoi de courriels -> Courriels sortants**

![eMailerOOo Spooler screenshot 1][52]

#### Liste des courriels sortants:

Chaque travaux d'envoi possède 3 états différents:
- État **0**: le courriel est prêt à être envoyé.
- État **1**: le courriel a été envoyé avec succès.
- État **2**: Une erreur est survenue lors de l'envoi du courriel. Vous pouvez consulter le message d'erreur dans le [Journal d'activité du spouleur][53].

![eMailerOOo Spooler screenshot 2][54]

Le spouleur de courriels est arrêté par défaut. **Il doit être démarré avec le bouton `Démarrer / Arrêter` pour que les courriels en attente soient envoyés**.

#### Journal d'activité du spouleur:

Lorsque le spouleur de courriel est démarré, son activité peut être visualisée dans le journal d'activité.

![eMailerOOo Spooler screenshot 3][55]

## A été testé avec:

* LibreOffice 7.3.7.2 - Lubuntu 22.04 - OpenJDK-11-JRE (amd64)

* LibreOffice 7.4.3.2(x64) - Windows 10(x64) - Adoptium JDK Hotspot 11.0.17 (x64) (sous Lubuntu 22.04 / VirtualBox 6.1.38)

* OpenOffice 4.1.13 - Lubuntu 22.04 - OpenJDK-11-JRE (amd64) (sous Lubuntu 22.04 / VirtualBox 6.1.38)

* **Ne fonctionne pas avec OpenOffice sous Windows** voir [dysfonctionnement 128569][11]. N'ayant aucune solution, je vous encourrage d'installer **LibreOffice**.

Je vous encourage en cas de problème :-(  
de créer un [dysfonctionnement][10]  
J'essaierai de le résoudre ;-)

## Historique:

### Ce qui a été fait pour la version 0.0.1:

- Ecriture de [IspDB][56] ou l'assistant de Configuration de connexion aux serveurs SMTP permettant:
    - De trouver les paramètres de connexion à un serveur SMTP à partir d'une adresse courriel. D'ailleur je remercie particulierement Mozilla, pour [Thunderbird autoconfiguration database][57] ou IspDB, qui à rendu ce défi possible...
    - D'afficher l'activité du service UNO `com.sun.star.mail.MailServiceProvider` lors de la connexion au serveur SMTP et l'envoi d'un courriel. 

- Ecriture du [Spouleur][58] de courriels permettant:
    - D'afficher les travaux d'envoi de courriel avec leurs états respectifs.
    - D'afficher l'activité du service UNO `com.sun.star.mail.SpoolerService` lors de l'envoi de courriels.
    - De démarrer et arrêter le service spouleur.

- Ecriture du [Merger][59] ou l'assistant de publipostage de courriels permettant:
    - De créer des listes de diffusions.
    - De fusionner et convertir au format HTML le document courant pour en faire le message du courriel.
    - De fusionner et/ou convertir au format PDF d'éventuel fichiers joints au courriel.

- Ecriture du [Mailer][60] de document permettant:
    - De convertir au format HTML le document pour en faire le message du courriel.
    - De convertir au format PDF d'éventuel fichiers joints au courriel.

- Ecriture d'un [Grid][61] piloté par un `com.sun.star.sdb.RowSet` permettant:
    - D'être paramètrable sur les colonnes à afficher.
    - D'être paramètrable sur l'ordre de tri à afficher.
    - De sauvegarder les paramètres d'affichage.

### Ce qui a été fait pour la version 0.0.2:

- Réécriture de [IspDB][56] ou Assistant de configuration de connexion aux serveurs de messagerie afin d'intégrer la configuration de la connexion IMAP.
  - Utilisation de [IMAPClient][62] version 2.2.0: une bibliothèque cliente IMAP complète, Pythonic et facile à utiliser.
  - Extension des fichiers IDL [com.sun.star.mail.*][63]:
    - [XMailMessage2.idl][64] prend désormais en charge la hiérarchisation des courriels (thread).
    - La nouvelle interface [XImapService][65] permet d'accéder à une partie de la bibliothèque IMAPClient.

- Réécriture du [Spouleur][66] afin d'intégrer des fonctionnalités IMAP comme la création d'un fil récapitulant le publipostage et regroupant tous les courriels envoyés.

- Soumission de l'extension eMailerOOo à Google et obtention de l'autorisation d'utiliser son API GMail afin d'envoyer des courriels avec un compte Google.

### Ce qui a été fait pour la version 0.0.3:

- Réécriture du [Grid][61] afin de permettre:
  - Le tri sur une colonne avec l'intégration du service UNO [SortableGridDataModel][67].
  - La génération des filtres des enregistrements nécessaires au service [Spouleur][58].
  - Le partage du module python avec l'extension [jdbcDriverOOo][68].

- Réécriture du [Merger][59] afin de permettre:
  - La gestion du nom du Schema dans de nom des tables afin d'être compatible avec la version 0.0.4 de [jdbcDriverOOo][69].
  - La création de liste de diffusion sur un groupe du carnet d'adresse et permettant de suivre la modification de son contenu.
  - L'utilisation de clé primaire, qui peuvent être composite, supportant les [DataType][70] `VARCHAR` et `INTEGER` ou derivé.
  - Un aperçu du document avec des champs de fusion remplis plus rapidement grâce au [Grid][61].

- Réécriture du [Spouleur][58] afin de permettre:
  - L'utilisation des nouveaux filtres supportant les clés primaires composite fourni par le [Merger][59].
  - L'utilisation du nouveau [Grid][61] permettant le tri sur une colonne.

- Encore plein d'autres choses...

### Que reste-t-il à faire pour la version 0.0.3:

- Ajouter de nouvelles langues pour l’internationalisation...

- Tout ce qui est bienvenu...

[1]: <img/eMailerOOo.svg>
[2]: <https://prrvchr.github.io/eMailerOOo/>
[3]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/TermsOfUse_fr>
[4]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/PrivacyPolicy_fr>
[5]: <https://prrvchr.github.io/eMailerOOo/README_fr#ce-qui-a-%C3%A9t%C3%A9-fait-pour-la-version-003>
[6]: <https://prrvchr.github.io/README_fr>
[7]: <https://fr.libreoffice.org/download/telecharger-libreoffice/>
[8]: <https://www.openoffice.org/fr/Telecharger/>
[9]: <https://github.com/prrvchr/eMailerOOo>
[10]: <https://github.com/prrvchr/eMailerOOo/issues/new>
[11]: <https://bz.apache.org/ooo/show_bug.cgi?id=128569>
[12]: <http://hsqldb.org/>
[13]: <https://wiki.documentfoundation.org/Documentation/HowTo/Install_the_correct_JRE_-_LibreOffice_on_Windows_10/fr>
[14]: <https://adoptium.net/releases.html?variant=openjdk11>
[15]: <https://bugs.documentfoundation.org/show_bug.cgi?id=139538>
[16]: <https://prrvchr.github.io/HsqlDBembeddedOOo/README_fr>
[17]: <https://prrvchr.github.io/OAuth2OOo/img/OAuth2OOo.svg>
[18]: <https://github.com/prrvchr/OAuth2OOo/raw/master/OAuth2OOo.oxt>
[19]: <https://prrvchr.github.io/jdbcDriverOOo/img/jdbcDriverOOo.svg>
[20]: <https://github.com/prrvchr/jdbcDriverOOo/raw/master/jdbcDriverOOo.oxt>
[21]: <https://prrvchr.github.io/vCardOOo/img/vCardOOo.svg>
[22]: <https://github.com/prrvchr/vCardOOo/raw/main/vCardOOo.oxt>
[23]: <https://fr.wikipedia.org/wiki/Nextcloud>
[24]: <https://prrvchr.github.io/gContactOOo/img/gContactOOo.svg>
[25]: <https://github.com/prrvchr/gContactOOo/raw/master/gContactOOo.oxt>
[26]: <https://prrvchr.github.io/mContactOOo/img/mContactOOo.svg>
[27]: <https://github.com/prrvchr/mContactOOo/raw/main/mContactOOo.oxt>
[28]: <https://github.com/prrvchr/eMailerOOo/raw/master/eMailerOOo.oxt>
[29]: <https://prrvchr.github.io/eMailerOOo/README_fr#publipostage-de-courriels-avec-des-listes-de-diffusion>
[30]: <https://prrvchr.github.io/eMailerOOo/README_fr#configuration-de-la-connexion>
[31]: <https://prrvchr.github.io/eMailerOOo/README_fr#courriels-sortants>
[32]: <img/eMailerOOo-Merger1_fr.png>
[33]: <img/eMailerOOo-Merger2_fr.png>
[34]: <img/eMailerOOo-Merger3_fr.png>
[35]: <img/eMailerOOo-Merger4_fr.png>
[36]: <https://prrvchr.github.io/eMailerOOo/README_fr#destinataires-disponibles>
[37]: <img/eMailerOOo-Merger5_fr.png>
[38]: <img/eMailerOOo-Merger6_fr.png>
[39]: <img/eMailerOOo-Merger7_fr.png>
[40]: <img/eMailerOOo-Merger8_fr.png>
[41]: <img/eMailerOOo-Merger9_fr.png>
[42]: <https://prrvchr.github.io/eMailerOOo/README_fr#configuration-de-la-connexion>
[43]: <img/eMailerOOo-Merger10_fr.png>
[44]: <img/eMailerOOo-Merger11_fr.png>
[45]: <https://prrvchr.github.io/eMailerOOo/README_fr#courriels-sortants>
[46]: <img/eMailerOOo-Ispdb1_fr.png>
[47]: <img/eMailerOOo-Ispdb2_fr.png>
[48]: <img/eMailerOOo-Ispdb3_fr.png>
[49]: <img/eMailerOOo-Ispdb4_fr.png>
[50]: <img/eMailerOOo-Ispdb5_fr.png>
[51]: <img/eMailerOOo-Ispdb6_fr.png>
[52]: <img/eMailerOOo-Spooler1_fr.png>
[53]: <https://prrvchr.github.io/eMailerOOo/README_fr#journal-dactivité-du-spouleur>
[54]: <img/eMailerOOo-Spooler2_fr.png>
[55]: <img/eMailerOOo-Spooler3_fr.png>
[56]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/smtpmailer/ispdb>
[57]: <https://wiki.mozilla.org/Thunderbird:Autoconfiguration>
[58]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/smtpmailer/spooler>
[59]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/smtpmailer/merger>
[60]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/smtpmailer/mailer>
[61]: <https://github.com/prrvchr/eMailerOOo/tree/master/uno/lib/uno/grid>
[62]: <https://github.com/mjs/imapclient#readme>
[63]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/idl/com/sun/star/mail>
[64]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailMessage2.idl>
[65]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XImapService.idl>
[66]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/smtpmailer/mailspooler.py>
[67]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/grid/SortableGridDataModel.html>
[68]: <https://github.com/prrvchr/jdbcDriverOOo/tree/master/source/jdbcDriverOOo/service/pythonpath/jdbcdriver/grid>
[69]: <https://prrvchr.github.io/jdbcDriverOOo/README_fr>
[70]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/sdbc/DataType.html>
