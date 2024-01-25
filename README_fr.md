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

**This [document][3] in english.**

**L'utilisation de ce logiciel vous soumet à nos [Conditions d'utilisation][4] et à notre [Politique de protection des données][5].**

# version [1.2.0][6]

## Introduction:

**eMailerOOo** fait partie d'une [Suite][7] d'extensions [LibreOffice][8] ~~et/ou [OpenOffice][9]~~ permettant de vous offrir des services inovants dans ces suites bureautique.

Cette extension vous permet d'envoyer des documents dans LibreOffice sous forme de courriel, éventuellement par publipostage, à vos contacts téléphoniques.

Elle fournit en plus une **API permettant d'envoyer des courriels en BASIC** et supportant les technologies les plus avancées (protocole OAuth2, Mozilla IspDB, HTTP au lieu de SMTP/IMAP, ...). Une macro [SendEmail][10] permettant d'envoyer des courriels est fournie à titre d'exemple.  
Si au préalable vous ouvrez un document, vous pouvez la lancer par:  
**Outils -> Macros -> Exécuter la macro... -> Mes macros -> eMailerOOo -> SendEmail -> Main -> Exécuter**

Etant un logiciel libre je vous encourage:
- A dupliquer son [code source][11].
- A apporter des modifications, des corrections, des améliorations.
- D'ouvrir un [dysfonctionnement][12] si nécessaire.

Bref, à participer au developpement de cette extension.  
Car c'est ensemble que nous pouvons rendre le Logiciel Libre plus intelligent.

___

## Prérequis:

L'extension eMailerOOo utilise l'extension OAuth2OOo pour fonctionner.  
Elle doit donc répondre aux [prérequis de l'extension OAuth2OOo][13].

L'extension eMailerOOo utilise l'extension jdbcDriverOOo pour fonctionner.  
Elle doit donc répondre aux [prérequis de l'extension jdbcDriverOOo][14].

**Sous Linux et macOS les paquets Python** utilisés par l'extension, peuvent s'il sont déja installé provenir du système et donc, **peuvent ne pas être à jour**.  
Afin de s'assurer que vos paquets Python sont à jour il est recommandé d'utiliser l'option **Info système** dans les Options de l'extension accessible par:  
**Outils -> Options -> Internet -> eMailerOOo -> Voir journal -> Info système**  
Si des paquets obsolètes apparaissent, vous pouvez les mettre à jour avec la commande:  
`pip install --upgrade <package-name>`

Pour plus d'information voir: [Ce qui a été fait pour la version 1.2.0][15].

___

## Installation:

Il semble important que le fichier n'ait pas été renommé lors de son téléchargement.  
Si nécessaire, renommez-le avant de l'installer.

- [![OAuth2OOo logo][17]][18] Installer l'extension **[OAuth2OOo.oxt][19]** [![Version][20]][19]

    Vous devez d'abord installer cette extension, si elle n'est pas déjà installée.

- [![jdbcDriverOOo logo][21]][22] Installer l'extension **[jdbcDriverOOo.oxt][23]** [![Version][24]][23]

    Cette extension est nécessaire pour utiliser HsqlDB version 2.7.2 avec toutes ses fonctionnalités.

- Si vous n'avez pas de source de données, vous pouvez:

    - [![vCardOOo logo][25]][26] Installer l'extension **[vCardOOo.oxt][27]** [![Version][28]][27]

        Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts présents sur une plateforme [**Nextcloud**][29] comme source de données pour les listes de diffusion et la fusion de documents.

    - [![gContactOOo logo][30]][31] Installer l'extension **[gContactOOo.oxt][32]** [![Version][33]][32]

        Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts téléphoniques personnels (contact Android) comme source de données pour les listes de diffusion et la fusion de documents.

    - [![mContactOOo logo][34]][35] Installer l'extension **[mContactOOo.oxt][36]** [![Version][37]][36]

        Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts Microsoft Outlook comme source de données pour les listes de diffusion et la fusion de documents.

- ![eMailerOOo logo][38] Installer l'extension **[eMailerOOo.oxt][39]** [![Version][40]][39]

Redémarrez LibreOffice après l'installation.  
**Attention, redémarrer LibreOffice peut ne pas suffire.**
- **Sous Windows** pour vous assurer que LibreOffice redémarre correctement, utilisez le Gestionnaire de tâche de Windows pour vérifier qu'aucun service LibreOffice n'est visible après l'arrêt de LibreOffice (et tuez-le si ç'est le cas).
- **Sous Linux ou macOS** vous pouvez également vous assurer que LibreOffice redémarre correctement, en le lançant depuis un terminal avec la commande `soffice` et en utilisant la combinaison de touches `Ctrl + C` si après l'arrêt de LibreOffice, le terminal n'est pas actif (pas d'invité de commande).

___

## Utilisation:

### Introduction:

Pour pouvoir utiliser la fonctionnalité de publipostage de courriels en utilisant des listes de diffusion, il est nécessaire d'avoir une **source de données** avec des tables ayant les colonnes suivantes:
- Une ou plusieurs colonnes d'adresses électroniques. Ces colonnes sont choisi dans une liste et si ce choix n'est pas unique, alors la première colonne d'adresse courriel non nulle sera utilisée.
- Une ou plusieurs colonnes de clé primaire permettant d'identifier de manière unique les enregistrements, elle peut être une clé primaire composée. Les types supportés sont VARCHAR et/ou INTEGER, ou derivé. Ces colonnes doivent être declarée avec la contrainte NOT NULL.

De plus, cette **source de données** doit avoir au moins une **table principale**, comprenant tous les enregistrements pouvant être utilisés lors du publipostage du courriel.

Si vous ne disposez pas d'une telle **source de données** alors je vous invite à installer une des extensions suivantes :
- [vCardOOo][26]. Cette extension vous permettra d'utiliser vos contacts présents sur une plateforme [**Nextcloud**][29] comme source de données.
- [gContactOOo][31]. Cette extension vous permettra d'utiliser votre téléphone Android (vos contacts téléphoniques) comme source de données.
- [mContactOOo][35]. Cette extension vous permettra d'utiliser vos contacts Microsoft Outlook comme source de données.

Pour ces 3 extensions le nom de la **table principale** peut être trouvé (et même changé avant toute connexion) dans:  
**Outils -> Options -> Internet -> Nom de l'extension -> Nom de la table principale**

Ce mode d'utilisation est composé de 3 sections:
- [Publipostage de courriels avec des listes de diffusion][41].
- [Configuration de la connexion][42].
- [Courriels sortants][43].

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

![eMailerOOo Merger screenshot 1][44]

#### Sélection de la source de données:

Le chargement de la source de données de l'assistant **Publipostage de courriels** devrait apparaître :

![eMailerOOo Merger screenshot 2][45]

Les captures d'écran suivantes utilisent l'extension [gContactOOo][31] comme **source de données**. Si vous utilisez votre propre **source de données**, il est nécessaire d'adapter les paramètres par rapport à celle-ci. 

Dans la copie d'écran suivante, on peut voir que la **source de données** gContactOOo s'appelle: `Adresses` et que dans la liste des tables la table: `PUBLIC.Tous mes contacts` est sélectionnée.

![eMailerOOo Merger screenshot 3][46]

Si aucune liste de diffusion n'existe, vous devez en créer une, en saisissant son nom et en validant avec: `ENTRÉE` ou le bouton `Ajouter`.

Assurez-vous lors de la création de la liste de diffusion que la **table principale** est toujours bien sélectionnée dans la liste des tables.  
Si cette recommandation n'est pas suivie alors **la fusion de documents ne fonctionnera pas** et ceci silencieusement.

![eMailerOOo Merger screenshot 4][47]

Maintenant que votre nouvelle liste de diffusion est disponible dans la liste, vous devez la sélectionner.

Et ajouter les colonnes suivantes:
- Colonne de clef primaire: `Uri`
- Colonnes d'adresses électronique: `HomeEmail`, `WorkEmail` et `OtherEmail`

Si plusieurs colonnes d'adresses courriel sont sélectionnées, alors l'ordre devient pertinent puisque le courriel sera envoyé à la première adresse disponible.  
De plus, à l'étape Sélection des destinataires de l'assistant, dans l'onglet [Destinataires disponibles][48], seuls les enregistrements avec au moins une colonne d'adresse courriel saisie seront répertoriés.  
Assurez-vous donc d'avoir un carnet d'adresses avec au moins un des champs d'adresse e-mail (Home, Work ou Other) renseigné.

![eMailerOOo Merger screenshot 5][49]

Ce paramètrage ne doit être effectué que pour les nouvelles listes de diffusion.  
Vous pouvez maintenant passer à l'étape suivante.

#### Sélection des destinataires:

##### Destinataires disponibles:

Les destinataires sont sélectionnés à l'aide de 2 boutons `Tout ajouter` et `Ajouter` permettant respectivement:
- Soit d'ajouter le groupe de destinataires sélectionnés dans la liste `Carnet d'adresses`. Ceci permet lors d'un publipostage, que les modifications du contenu du groupe soient prises en compte. Une liste de diffusion n'accepte qu'un seul groupe.
- Soit d'ajouter la sélection, qui peut être multiple à l'aide de la touche `CTRL`. Cette sélection est immuable quelle que soit la modification des groupes du carnet d'adresses.

![eMailerOOo Merger screenshot 6][50]

Example de la sélection multiple:

![eMailerOOo Merger screenshot 7][51]

##### Destinataires sélectionnés:

Les destinataires sont désélectionnés à l'aide de 2 boutons `Tout retirer` et `Retirer` permettant respectivement:
- Soit de retirer le groupe qui a été affecté à cette liste de diffusion. Ceci est nécessaire afin de pouvoir modifier à nouveau le contenu de cette liste de diffusion.
- Soit de retirer la sélection, qui peut être multiple à l'aide de la touche `CTRL`.

![eMailerOOo Merger screenshot 8][52]

Si vous avez sélectionné au moins 1 destinataire, vous pouvez passer à l'étape suivante.

#### Sélection des options d'envoi:

Si cela n'est pas déjà fait, vous devez créer un nouvel expéditeur à l'aide du bouton `Ajouter`.

![eMailerOOo Merger screenshot 9][53]

La création du nouvel expéditeur est décrite dans la section [Configuration de la connexion][42].

Le courriel doit avoir un sujet. Il peut être enregistré dans le document Writer.  
Vous pouvez insérer des champs de fusion dans l'objet du courriel. Un champ de fusion est composé d'une accolade ouvrante, du nom de la colonne référencée (sensible à la casse) et d'une accolade fermante (ie: `{NomColonne}`).

![eMailerOOo Merger screenshot 10][54]

Le courriel peut éventuellement contenir des fichiers joints. Ils peuvent être enregistrés dans le document Writer.  
La capture d'écran suivante montre 1 fichier joint qui sera fusionné sur la source de données puis converti au format PDF avant d'être joint au courriel.

![eMailerOOo Merger screenshot 11][55]

Assurez-vous de toujours quitter l'assistant avec le bouton `Terminer` pour confirmer la soumission des travaux d'envoi.  
Pour envoyer les travaux d'envoi, veuillez suivre la section [Courriels sortants][43].

### Configuration de la connexion:

#### Démarrage de l'assistant de connexion:

Dans LibreOffice / OpenOffice aller à: **Outils -> Add-ons -> Envoi de courriels -> Configurer la connexion**

![eMailerOOo Ispdb screenshot 1][56]

#### Sélection du compte:

![eMailerOOo Ispdb screenshot 2][57]

#### Trouver la configuration:

![eMailerOOo Ispdb screenshot 3][58]

#### Configuration SMTP:

![eMailerOOo Ispdb screenshot 4][59]

#### Configuration IMAP:

![eMailerOOo Ispdb screenshot 5][60]

#### Tester la connexion:

![eMailerOOo Ispdb screenshot 6][61]

Assurez-vous de toujours quitter l'assistant avec le bouton `Terminer` afin d'enregistrer les paramètres de connexion.

### Courriels sortants:

#### Démarrage du spouleur de courriels:

Dans LibreOffice / OpenOffice aller à: **Outils -> Add-ons -> Envoi de courriels -> Courriels sortants**

![eMailerOOo Spooler screenshot 1][62]

#### Liste des courriels sortants:

Chaque travaux d'envoi possède 3 états différents:
- État **0**: le courriel est prêt à être envoyé.
- État **1**: le courriel a été envoyé avec succès.
- État **2**: Une erreur est survenue lors de l'envoi du courriel. Vous pouvez consulter le message d'erreur dans le [Journal d'activité du spouleur][63].

![eMailerOOo Spooler screenshot 2][64]

Le spouleur de courriels est arrêté par défaut. **Il doit être démarré avec le bouton `Démarrer / Arrêter` pour que les courriels en attente soient envoyés**.

#### Journal d'activité du spouleur:

Lorsque le spouleur de courriel est démarré, son activité peut être visualisée dans le journal d'activité.

![eMailerOOo Spooler screenshot 3][65]

___

## A été testé avec:

* LibreOffice 7.3.7.2 - Lubuntu 22.04 - Python version 3.10.12 - OpenJDK-11-JRE (amd64)

* LibreOffice 7.5.4.2(x86) - Windows 10 - Python version 3.8.16 - Adoptium JDK Hotspot 11.0.19 (under Lubuntu 22.04 / VirtualBox 6.1.38)

* LibreOffice 7.4.3.2(x64) - Windows 10(x64) - Python version 3.8.15  - Adoptium JDK Hotspot 11.0.17 (x64) (under Lubuntu 22.04 / VirtualBox 6.1.38)

* **Ne fonctionne pas avec OpenOffice sous Windows** voir [dysfonctionnement 128569][66]. N'ayant aucune solution, je vous encourrage d'installer **LibreOffice**.

Je vous encourage en cas de problème :confused:  
de créer un [dysfonctionnement][11]  
J'essaierai de le résoudre :smile:

___

## Historique:

### Ce qui a été fait pour la version 0.0.1:

- Ecriture de [IspDB][67] ou l'assistant de Configuration de connexion aux serveurs SMTP permettant:
    - De trouver les paramètres de connexion à un serveur SMTP à partir d'une adresse courriel. D'ailleur je remercie particulierement Mozilla, pour [Thunderbird autoconfiguration database][68] ou IspDB, qui à rendu ce défi possible...
    - D'afficher l'activité du service UNO `com.sun.star.mail.MailServiceProvider` lors de la connexion au serveur SMTP et l'envoi d'un courriel. 

- Ecriture du [Spouleur][69] de courriels permettant:
    - D'afficher les travaux d'envoi de courriel avec leurs états respectifs.
    - D'afficher l'activité du service UNO `com.sun.star.mail.SpoolerService` lors de l'envoi de courriels.
    - De démarrer et arrêter le service spouleur.

- Ecriture du [Merger][70] ou l'assistant de publipostage de courriels permettant:
    - De créer des listes de diffusions.
    - De fusionner et convertir au format HTML le document courant pour en faire le message du courriel.
    - De fusionner et/ou convertir au format PDF d'éventuel fichiers joints au courriel.

- Ecriture du [Mailer][71] de document permettant:
    - De convertir au format HTML le document pour en faire le message du courriel.
    - De convertir au format PDF d'éventuel fichiers joints au courriel.

- Ecriture d'un [Grid][72] piloté par un `com.sun.star.sdb.RowSet` permettant:
    - D'être paramètrable sur les colonnes à afficher.
    - D'être paramètrable sur l'ordre de tri à afficher.
    - De sauvegarder les paramètres d'affichage.

### Ce qui a été fait pour la version 0.0.2:

- Réécriture de [IspDB][67] ou Assistant de configuration de connexion aux serveurs de messagerie afin d'intégrer la configuration de la connexion IMAP.
    - Utilisation de [IMAPClient][73] version 2.2.0: une bibliothèque cliente IMAP complète, Pythonic et facile à utiliser.
    - Extension des fichiers IDL [com.sun.star.mail.*][74]:
        - [XMailMessage2.idl][75] prend désormais en charge la hiérarchisation des courriels (thread).
        - La nouvelle interface [XImapService][76] permet d'accéder à une partie de la bibliothèque IMAPClient.

- Réécriture du [Spouleur][77] afin d'intégrer des fonctionnalités IMAP comme la création d'un fil récapitulant le publipostage et regroupant tous les courriels envoyés.

- Soumission de l'extension eMailerOOo à Google et obtention de l'autorisation d'utiliser son API GMail afin d'envoyer des courriels avec un compte Google.

### Ce qui a été fait pour la version 0.0.3:

- Réécriture du [Grid][72] afin de permettre:
    - Le tri sur une colonne avec l'intégration du service UNO [SortableGridDataModel][78].
    - La génération des filtres des enregistrements nécessaires au service [Spouleur][69].
    - Le partage avec le module python [Grid][79] de l'extension [jdbcDriverOOo][22].

- Réécriture du [Merger][70] afin de permettre:
    - La gestion du nom du Schema dans de nom des tables afin d'être compatible avec la version 0.0.4 de [jdbcDriverOOo][22].
    - La création de liste de diffusion sur un groupe du carnet d'adresse et permettant de suivre la modification de son contenu.
    - L'utilisation de clé primaire, qui peuvent être composite, supportant les [DataType][80] `VARCHAR` et `INTEGER` ou derivé.
    - Un aperçu du document avec des champs de fusion remplis plus rapidement grâce au [Grid][72].

- Réécriture du [Spouleur][69] afin de permettre:
    - L'utilisation des nouveaux filtres supportant les clés primaires composite fourni par le [Merger][70].
    - L'utilisation du nouveau [Grid][72] permettant le tri sur une colonne.

- Encore plein d'autres choses...

### Ce qui a été fait pour la version 1.0.0:

- L'extension **smtpMailerOOo** a été renomé en **eMailerOOo**.

### Ce qui a été fait pour la version 1.0.1:

- L'absence ou l'obsolescence des extensions **OAuth2OOo** et/ou **jdbcDriverOOo** nécessaires au bon fonctionnement de **eMailerOOo** affiche désormais un message d'erreur. Ceci afin d'éviter qu'un dysfonctionnement tel que le [dysfonctionnement #3][81] ne se reproduise...

- La base de données HsqlDB sous-jacente peut être ouverte dans Base avec: **Outils -> Options -> Internet -> eMailerOOo -> Base de données**.

- Le menu **Outils -> Add-ons** s'affiche désormais correctement en fonction du contexte.

- Encore plein d'autres choses...

### Ce qui a été fait pour la version 1.0.2:

- Si aucune configuration n'est trouvée dans l'assistant de configuration de la connexion (IspDB Wizard) alors il est possible de configurer la connexion manuellement. Voir [dysfonctionnement #5][82].

### Ce qui a été fait pour la version 1.1.0:

- Dans l'assistant de configuration de la connexion (IspDB Wizard) il est maintenant possible de désactiver la configuration IMAP.  
    En conséquence, cela n'envoie plus de fil de discussion (message IMAP) lors de la fusion d'un mailing.  
    Dans ce même assistant, il est désormais possible de saisir une adresse courriel de réponse.

- Dans l'assistant de fusion d'email, il est désormais possible d'insérer des champs de fusion dans l'objet du courriel. Voir [dysfonctionnement #6][83].  
    Dans le sujet d'un courriel, un champ de fusion est composé d'une accolade ouvrante, du nom de la colonne référencée (sensible à la casse) et d'une accolade fermante (ie: `{NomDeLaColonne}`).  
    Lors de la saisie du sujet du courriel, une erreur de syntaxe dans un champ de fusion sera signalée et empêchera la soumission du mailing.

- Il est désormais possible dans le Spooler de visualiser les courriels au format eml.

- Un service [com.sun.star.mail.MailUser][84] permet désormais d'accéder à une configuration de connexion (SMTP et/ou IMAP) depuis une adresse courriel qui suite la rfc822.  
    Un autre service [com.sun.star.datatransfer.TransferableFactory][85] permet, comme son nom l'indique, la création de [Transferable][86] à partir d'un texte (string), d'une séquence binaire, d'une Url (file://...) ou un flux de données (InputStream).  
    Ces deux nouveaux services simplifient grandement l'API mail de LibreOffice et permettent d'envoyer des courriels depuis Basic. Voir le [dysfonctionnement #4][87].  
    Vous trouverez une macro Basic vous permettant d'envoyer des emails dans : **Outils -> Macros -> Editer les Macros... -> eMailerOOo -> SendEmail**.

### Ce qui a été fait pour la version 1.1.1:

- Prise en charge de la version 1.2.0 de l'extension **OAuth2OOo**. Les versions précédentes ne fonctionneront pas avec l'extension **OAuth2OOo** 1.2.0 ou ultérieure.

### Ce qui a été fait pour la version 1.2.0:

- Tous les paquets Python nécessaires à l'extension sont désormais enregistrés dans un fichier [requirements.txt][88] suivant la [PEP 508][89].
- Désormais si vous n'êtes pas sous Windows alors les paquets Python nécessaires à l'extension peuvent être facilement installés avec la commande:  
  `pip install requirements.txt`
- Modification de la section [Prérequis][90].

### Que reste-t-il à faire pour la version 1.2.0:

- Ajouter de nouvelles langues pour l’internationalisation...

- Tout ce qui est bienvenu...

[1]: </img/emailer.svg#collapse>
[2]: <https://prrvchr.github.io/eMailerOOo/>
[3]: <https://prrvchr.github.io/eMailerOOo/>
[4]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/TermsOfUse_fr>
[5]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/PrivacyPolicy_fr>
[6]: <https://prrvchr.github.io/eMailerOOo/README_fr#ce-qui-a-%C3%A9t%C3%A9-fait-pour-la-version-110>
[7]: <https://prrvchr.github.io/>
[8]: <https://www.libreoffice.org/download/download-libreoffice/>
[9]: <https://www.openoffice.org/download/index.html>
[10]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/eMailerOOo/SendEmail.xba>
[11]: <https://github.com/prrvchr/eMailerOOo>
[12]: <https://github.com/prrvchr/eMailerOOo/issues/new>
[13]: <https://prrvchr.github.io/OAuth2OOo/README_fr#pr%C3%A9requis>
[14]: <https://prrvchr.github.io/jdbcDriverOOo/README_fr#pr%C3%A9requis>
[15]: <https://prrvchr.github.io/eMailerOOo/README_fr#ce-qui-a-%C3%A9t%C3%A9-fait-pour-la-version-120>
[17]: <https://prrvchr.github.io/OAuth2OOo/img/OAuth2OOo.svg#middle>
[18]: <https://prrvchr.github.io/OAuth2OOo/README_fr>
[19]: <https://github.com/prrvchr/OAuth2OOo/releases/latest/download/OAuth2OOo.oxt>
[20]: <https://img.shields.io/github/v/tag/prrvchr/OAuth2OOo?label=latest#right>
[21]: <https://prrvchr.github.io/jdbcDriverOOo/img/jdbcDriverOOo.svg#middle>
[22]: <https://prrvchr.github.io/jdbcDriverOOo/README_fr>
[23]: <https://github.com/prrvchr/jdbcDriverOOo/releases/latest/download/jdbcDriverOOo.oxt>
[24]: <https://img.shields.io/github/v/tag/prrvchr/jdbcDriverOOo?label=latest#right>
[25]: <https://prrvchr.github.io/vCardOOo/img/vCardOOo.svg#middle>
[26]: <https://prrvchr.github.io/vCardOOo/README_fr>
[27]: <https://github.com/prrvchr/vCardOOo/releases/latest/download/vCardOOo.oxt>
[28]: <https://img.shields.io/github/v/tag/prrvchr/vCardOOo?label=latest#right>
[29]: <https://fr.wikipedia.org/wiki/Nextcloud>
[30]: <https://prrvchr.github.io/gContactOOo/img/gContactOOo.svg#middle>
[31]: <https://prrvchr.github.io/gContactOOo/README_fr>
[32]: <https://github.com/prrvchr/gContactOOo/releases/latest/download/gContactOOo.oxt>
[33]: <https://img.shields.io/github/v/tag/prrvchr/gContactOOo?label=latest#right>
[34]: <https://prrvchr.github.io/mContactOOo/img/mContactOOo.svg#middle>
[35]: <https://prrvchr.github.io/mContactOOo/README_fr>
[36]: <https://github.com/prrvchr/mContactOOo/releases/latest/download/mContactOOo.oxt>
[37]: <https://img.shields.io/github/v/tag/prrvchr/mContactOOo?label=latest#right>
[38]: <https://prrvchr.github.io/eMailerOOo/img/eMailerOOo.svg#middle>
[39]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/eMailerOOo.oxt>
[40]: <https://img.shields.io/github/downloads/prrvchr/eMailerOOo/latest/total?label=v1.2.0#right>
[41]: <https://prrvchr.github.io/eMailerOOo/README_fr#publipostage-de-courriels-avec-des-listes-de-diffusion>
[42]: <https://prrvchr.github.io/eMailerOOo/README_fr#configuration-de-la-connexion>
[43]: <https://prrvchr.github.io/eMailerOOo/README_fr#courriels-sortants>
[44]: <img/eMailerOOo-Merger1.png>
[45]: <img/eMailerOOo-Merger2.png>
[46]: <img/eMailerOOo-Merger3.png>
[47]: <img/eMailerOOo-Merger4.png>
[48]: <https://prrvchr.github.io/eMailerOOo/README_fr#destinataires-disponibles>
[49]: <img/eMailerOOo-Merger5.png>
[50]: <img/eMailerOOo-Merger6.png>
[51]: <img/eMailerOOo-Merger7.png>
[52]: <img/eMailerOOo-Merger8.png>
[53]: <img/eMailerOOo-Merger9.png>
[54]: <img/eMailerOOo-Merger10.png>
[55]: <img/eMailerOOo-Merger11.png>
[56]: <img/eMailerOOo-Ispdb1.png>
[57]: <img/eMailerOOo-Ispdb2.png>
[58]: <img/eMailerOOo-Ispdb3.png>
[59]: <img/eMailerOOo-Ispdb4.png>
[60]: <img/eMailerOOo-Ispdb5.png>
[61]: <img/eMailerOOo-Ispdb6.png>
[62]: <img/eMailerOOo-Spooler1.png>
[63]: <https://prrvchr.github.io/eMailerOOo/README_fr#journal-dactivité-du-spouleur>
[64]: <img/eMailerOOo-Spooler2.png>
[65]: <img/eMailerOOo-Spooler3.png>
[66]: <https://bz.apache.org/ooo/show_bug.cgi?id=128569>
[67]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/ispdb>
[68]: <https://wiki.mozilla.org/Thunderbird:Autoconfiguration>
[69]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler>
[70]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/merger>
[71]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/mailer>
[72]: <https://github.com/prrvchr/eMailerOOo/tree/master/uno/lib/uno/grid>
[73]: <https://github.com/mjs/imapclient#readme>
[74]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/idl/com/sun/star/mail>
[75]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailMessage2.idl>
[76]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XImapService.idl>
[77]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/spooler.py>
[78]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/grid/SortableGridDataModel.html>
[79]: <https://github.com/prrvchr/jdbcDriverOOo/tree/master/source/jdbcDriverOOo/service/pythonpath/jdbcdriver/grid>
[80]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/sdbc/DataType.html>
[81]: <https://github.com/prrvchr/eMailerOOo/issues/3>
[82]: <https://github.com/prrvchr/eMailerOOo/issues/5>
[83]: <https://github.com/prrvchr/eMailerOOo/issues/6>
[84]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailUser.idl>
[85]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/datatransfer/XTransferableFactory.idl>
[86]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/datatransfer/XTransferable.html>
[87]: <https://github.com/prrvchr/eMailerOOo/issues/4>
[88]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/requirements.txt>
[89]: <https://peps.python.org/pep-0508/>
[90]: <https://prrvchr.github.io/eMailerOOo/README_fr#pr%C3%A9requis>
