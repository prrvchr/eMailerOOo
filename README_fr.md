---
layout: default
title: eMailerOOo documentation (Français)
permalink: /fr/
redirect_from:
  - /README_fr
  - /README_fr.html
---
<!--
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-25 https://prrvchr.github.io                                  ║
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
# [![eMailerOOo logo][1]][2] Documentation

**This [document][3] in english.**

**L'utilisation de ce logiciel vous soumet à nos [Conditions d'Utilisation][4] et à notre [Politique de Protection des Données][5].**

# version [1.5.2][6]

## Introduction:

**eMailerOOo** fait partie d'une [Suite][7] d'extensions [LibreOffice][8] ~~et/ou [OpenOffice][9]~~ permettant de vous offrir des services inovants dans ces suites bureautique.

Cette extension vous permet d'envoyer des documents dans LibreOffice sous forme de courriel, éventuellement par publipostage, à vos contacts téléphoniques.

Elle fournit en plus une API permettant d'**[envoyer des courriels en BASIC][10]** et supportant les technologies les plus avancées: protocole OAuth2, Mozilla IspDB, HTTP au lieu de SMTP/IMAP pour les serveurs Google...  
Une macro **SendEmail** permettant d'envoyer des courriels est fournie à titre d'exemple. Si au préalable vous ouvrez un document, vous pouvez la lancer par:  
**Outils -> Macros -> Exécuter la macro... -> Mes macros -> eMailerOOo -> SendEmail -> Main -> Exécuter**

Etant un logiciel libre je vous encourage:
- A dupliquer son [code source][11].
- A apporter des modifications, des corrections, des améliorations.
- D'ouvrir un [dysfonctionnement][12] si nécessaire.
- De [participer au frais][13] de la [certification CASA][14].

Bref, à participer au developpement de cette extension.  
Car c'est ensemble que nous pouvons rendre le Logiciel Libre plus intelligent.

___

## Certification CASA:

Afin de garantir l'interopérabilité avec **Google**, l'extension **eMailerOOo** utilise l'extension **OAuth2OOo** qui nécessite la [certification CASA][14].  
Jusqu'à présent, cette certification était gratuite et réalisée par un partenaire de Google.  
L'application **OAuth2OOo** a obtenu sa [certification CASA][15] le 28/11/2023.

**Maintenant cette certification est devenue désormais payante et coûte 600$.**

Je n'avais jamais anticipé de tels frais et je compte sur votre contribution pour financer cette certification.

Merci pour votre aide. [![Sponsor][16]][13]

___

## Prérequis:

L'extension eMailerOOo utilise l'extension OAuth2OOo pour fonctionner.  
Elle doit donc répondre aux [prérequis de l'extension OAuth2OOo][17].

L'extension eMailerOOo utilise l'extension jdbcDriverOOo pour fonctionner.  
Elle doit donc répondre aux [prérequis de l'extension jdbcDriverOOo][18].  
De plus, eMailerOOo nécessite que l'extension jdbcDriverOOo soit configurée pour fournir `com.sun.star.sdb` comme niveau d'API, qui est la configuration par défaut.

___

## Installation:

Il semble important que le fichier n'ait pas été renommé lors de son téléchargement.  
Si nécessaire, renommez-le avant de l'installer.

- [![OAuth2OOo logo][19]][20] Installer l'extension **[OAuth2OOo.oxt][21]** [![Version][22]][21]

    Vous devez d'abord installer cette extension, si elle n'est pas déjà installée.

- [![jdbcDriverOOo logo][23]][24] Installer l'extension **[jdbcDriverOOo.oxt][25]** [![Version][26]][25]

    Cette extension est nécessaire pour utiliser HsqlDB version 2.7.2 avec toutes ses fonctionnalités.

- Si vous n'avez pas de source de données, vous pouvez:

    - [![vCardOOo logo][27]][28] Installer l'extension **[vCardOOo.oxt][29]** [![Version][30]][29]

        Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts présents sur une plateforme [Nextcloud][31] comme source de données pour les listes de diffusion et la fusion de documents.

    - [![gContactOOo logo][32]][33] Installer l'extension **[gContactOOo.oxt][34]** [![Version][35]][34]

        Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts téléphoniques personnels (contact Android) comme source de données pour les listes de diffusion et la fusion de documents.

    - [![mContactOOo logo][36]][37] Installer l'extension **[mContactOOo.oxt][38]** [![Version][39]][38]

        Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts Microsoft Outlook comme source de données pour les listes de diffusion et la fusion de documents.

    - [![HyperSQLOOo logo][40]][41] Installer l'extension **[HyperSQLOOo.oxt][42]** [![Version][43]][42]

        Cette extension n'est nécessaire que si vous souhaitez utiliser un fichier Calc comme source de données pour les listes de diffusion et la fusion de documents. Voir: [Comment importer des données depuis un fichier Calc][44].

- ![eMailerOOo logo][45] Installer l'extension **[eMailerOOo.oxt][46]** [![Version][47]][46]

Redémarrez LibreOffice après l'installation.  
**Attention, redémarrer LibreOffice peut ne pas suffire.**
- **Sous Windows** pour vous assurer que LibreOffice redémarre correctement, utilisez le Gestionnaire de tâche de Windows pour vérifier qu'aucun service LibreOffice n'est visible après l'arrêt de LibreOffice (et tuez-le si ç'est le cas).
- **Sous Linux ou macOS** vous pouvez également vous assurer que LibreOffice redémarre correctement, en le lançant depuis un terminal avec la commande `soffice` et en utilisant la combinaison de touches `Ctrl + C` si après l'arrêt de LibreOffice, le terminal n'est pas actif (pas d'invité de commande).

Après ce redémarrage, il vous sera demandé **d'installer des paquets Python contenant des fichiers binaires**. Veuillez consulter la section [Installation de paquets Python][48] pour plus d'informations.

___

## Utilisation:

### Introduction:

Pour pouvoir utiliser la fonctionnalité de publipostage de courriels en utilisant des listes de diffusion, il est nécessaire d'avoir une **source de données** avec des tables ayant les colonnes suivantes:
- Une ou plusieurs colonnes d'adresses électroniques. Ces colonnes sont choisi dans une liste et si ce choix n'est pas unique, alors la première colonne d'adresse courriel non nulle sera utilisée.
- Une ou plusieurs colonnes de clé primaire permettant d'identifier de manière unique les enregistrements, elle peut être une clé primaire composée. Les types supportés sont VARCHAR et/ou INTEGER, ou derivé. Ces colonnes doivent être declarée avec la contrainte NOT NULL.

De plus, cette **source de données** doit avoir au moins une **table principale**, comprenant tous les enregistrements pouvant être utilisés lors du publipostage du courriel.

Si vous ne disposez pas d'une telle **source de données** alors je vous invite à installer une des extensions suivantes :
- [vCardOOo][28]. Cette extension vous permettra d'utiliser vos contacts présents sur une plateforme [Nextcloud][31] comme source de données.
- [gContactOOo][33]. Cette extension vous permettra d'utiliser votre téléphone Android (vos contacts téléphoniques) comme source de données.
- [mContactOOo][37]. Cette extension vous permettra d'utiliser vos contacts Microsoft Outlook comme source de données.
- [HyperSQLOOo][41]. Cette extension vous permettra d'utiliser un fichier Calc comme source de données. Voir: [Comment importer des données depuis un fichier Calc][44].

Pour les 3 premières extensions le nom de la **table principale** peut être trouvé (et même changé avant toute connexion) dans:  
**Outils -> Options -> Internet -> Nom de l'extension -> Nom de la table principale**

Ce mode d'utilisation est composé de 3 sections:
- [Publipostage de courriels avec des listes de diffusion][49].
- [Configuration de la connexion][50].
- [Courriels sortants][51].

### Publipostage de courriels avec des listes de diffusion:

#### Prérequis:

Pour pouvoir publiposter des courriels suivant une liste de diffusion, vous devez:
- Disposer d'une **source de données** comme décrit dans l'introduction précédente.
- Ouvrir un **nouveau document** Writer dans LibreOffice.

Ce document Writer peut inclure des champs de fusion (insérables par la commande: **Insertion -> Champ -> Autres champs -> Base de données -> Champ de publipostage**), cela est même nécessaire si vous souhaitez pouvoir personnaliser le contenu du courriel et d'eventuel fichiers attachés.  
Ces champs de fusion doivent uniquement faire référence à la **table principale** de la **source de données**.

Si vous souhaitez utiliser un **document Writer déja existant**, vous devez vous assurer en plus que la **source de données** et la **table principale** sont bien rattachées au document dans : **Outils -> Source du carnet d'adresses...**.

Si ces recommandations ne sont pas suivies alors **la fusion de documents ne fonctionnera pas** et ceci silencieusement.

#### Démarrage de l'assistant de publipostage de courriels:

Dans un document LibreOffice Writer aller à: **Outils -> Add-ons -> Envoi de courriels -> Publiposter un document**

![eMailerOOo Merger screenshot 1][52]

#### Sélection de la source de données:

Le chargement de la source de données de l'assistant **Publipostage de courriels** devrait apparaître :

![eMailerOOo Merger screenshot 2][53]

Les captures d'écran suivantes utilisent l'extension [gContactOOo][33] comme **source de données**. Si vous utilisez votre propre **source de données**, il est nécessaire d'adapter les paramètres par rapport à celle-ci. 

Dans la copie d'écran suivante, on peut voir que la **source de données** gContactOOo s'appelle: `Adresses` et que dans la liste des tables la table: `PUBLIC.Tous mes contacts` est sélectionnée.

![eMailerOOo Merger screenshot 3][54]

Si aucune liste de diffusion n'existe, vous devez en créer une, en saisissant son nom et en validant avec: `ENTRÉE` ou le bouton `Ajouter`.

Assurez-vous lors de la création de la liste de diffusion que la **table principale** est toujours bien sélectionnée dans la liste des tables.  
Si cette recommandation n'est pas suivie alors **la fusion de documents ne fonctionnera pas** et ceci silencieusement.

![eMailerOOo Merger screenshot 4][55]

Maintenant que votre nouvelle liste de diffusion est disponible dans la liste, vous devez la sélectionner.

Et ajouter les colonnes suivantes:
- Colonne de clef primaire: `Uri`
- Colonnes d'adresses électronique: `HomeEmail`, `WorkEmail` et `OtherEmail`

Si plusieurs colonnes d'adresses courriel sont sélectionnées, alors l'ordre devient pertinent puisque le courriel sera envoyé à la première adresse disponible.  
De plus, à l'étape Sélection des destinataires de l'assistant, dans l'onglet [Destinataires disponibles][56], seuls les enregistrements avec au moins une colonne d'adresse courriel saisie seront répertoriés.  
Assurez-vous donc d'avoir un carnet d'adresses avec au moins un des champs d'adresse e-mail (Home, Work ou Other) renseigné.

![eMailerOOo Merger screenshot 5][57]

Ce paramètrage ne doit être effectué que pour les nouvelles listes de diffusion.  
Vous pouvez maintenant passer à l'étape suivante.

#### Sélection des destinataires:

##### Destinataires disponibles:

Les destinataires sont sélectionnés à l'aide de 2 boutons `Tout ajouter` et `Ajouter` permettant respectivement:
- Soit d'ajouter le groupe de destinataires sélectionnés dans la liste `Carnet d'adresses`. Ceci permet lors d'un publipostage, que les modifications du contenu du groupe soient prises en compte. Une liste de diffusion n'accepte qu'un seul groupe.
- Soit d'ajouter la sélection, qui peut être multiple à l'aide de la touche `CTRL`. Cette sélection est immuable quelle que soit la modification des groupes du carnet d'adresses.

![eMailerOOo Merger screenshot 6][58]

Example de la sélection multiple:

![eMailerOOo Merger screenshot 7][59]

##### Destinataires sélectionnés:

Les destinataires sont désélectionnés à l'aide de 2 boutons `Tout retirer` et `Retirer` permettant respectivement:
- Soit de retirer le groupe qui a été affecté à cette liste de diffusion. Ceci est nécessaire afin de pouvoir modifier à nouveau le contenu de cette liste de diffusion.
- Soit de retirer la sélection, qui peut être multiple à l'aide de la touche `CTRL`.

![eMailerOOo Merger screenshot 8][60]

Si vous avez sélectionné au moins 1 destinataire, vous pouvez passer à l'étape suivante.

#### Sélection des options d'envoi:

Si cela n'est pas déjà fait, vous devez créer un nouvel expéditeur à l'aide du bouton `Ajouter`.

![eMailerOOo Merger screenshot 9][61]

La création du nouvel expéditeur est décrite dans la section [Configuration de la connexion][50].

Le courriel doit avoir un sujet. Il peut être enregistré dans le document Writer.  
Vous pouvez insérer des champs de fusion dans l'objet du courriel. Un champ de fusion est composé d'une accolade ouvrante, du nom de la colonne référencée (sensible à la casse) et d'une accolade fermante (ie: `{NomColonne}`).

![eMailerOOo Merger screenshot 10][62]

Le courriel peut éventuellement contenir des fichiers joints. Ils peuvent être enregistrés dans le document Writer.  
La capture d'écran suivante montre 1 fichier joint qui sera fusionné sur la source de données puis converti au format PDF avant d'être joint au courriel.

![eMailerOOo Merger screenshot 11][63]

Assurez-vous de toujours quitter l'assistant avec le bouton `Terminer` pour confirmer la soumission des travaux d'envoi.  
Pour envoyer les travaux d'envoi, veuillez suivre la section [Courriels sortants][51].

### Configuration de la connexion:

#### Démarrage de l'assistant de connexion:

Dans LibreOffice aller à: **Outils -> Add-ons -> Envoi de courriels -> Configurer la connexion**

![eMailerOOo Ispdb screenshot 1][64]

#### Sélection du compte:

![eMailerOOo Ispdb screenshot 2][65]

#### Trouver la configuration:

![eMailerOOo Ispdb screenshot 3][66]

#### Configuration SMTP:

![eMailerOOo Ispdb screenshot 4][67]

#### Configuration IMAP:

![eMailerOOo Ispdb screenshot 5][68]

#### Tester la connexion:

![eMailerOOo Ispdb screenshot 6][69]

Assurez-vous de toujours quitter l'assistant avec le bouton `Terminer` afin d'enregistrer les paramètres de connexion.

### Courriels sortants:

#### Démarrage du spouleur de courriels:

Dans LibreOffice aller à: **Outils -> Add-ons -> Envoi de courriels -> Courriels sortants**

![eMailerOOo Spooler screenshot 1][70]

#### Liste des courriels sortants:

Chaque travaux d'envoi possède 3 états différents:
- État **0**: le courriel est prêt à être envoyé.
- État **1**: le courriel a été envoyé avec succès.
- État **2**: Une erreur est survenue lors de l'envoi du courriel. Vous pouvez consulter le message d'erreur dans le [Journal d'activité du spouleur][71].

![eMailerOOo Spooler screenshot 2][72]

Le spouleur de courriels est arrêté par défaut. **Il doit être démarré avec le bouton `Démarrer / Arrêter` pour que les courriels en attente soient envoyés**.

#### Journal d'activité du spouleur:

Lorsque le spouleur de courriel est démarré, son activité peut être visualisée dans le journal d'activité.

![eMailerOOo Spooler screenshot 3][73]

___

## Envoi de courriel avec une macro LibreOffice en Basic:

Il est possible d'envoyer des courriels à l'aide de **macros écrites en Basic**. L'envoi d'un courriel nécessite une macro de quelques 50 lignes de code et pourra supporter la plupart des serveurs SMTP/IMAP.  
Voici le code minimum nécessaire pour envoyer un courriel avec des fichiers joints.

```
Sub Main

    Rem Demandez à l’utilisateur une adresse courriel d’expéditeur.
    sSender = InputBox("Veuillez saisir l'adresse courriel de l'expéditeur")
    Rem L'utilisateur a cliqué sur Annuler.
    if sSender = "" then
        exit sub
    endif

    Rem Demandez à l'utilisateur l'adresse courriel du destinataire.
    sRecipient = InputBox("Veuillez saisir l'adresse courriel du destinataire")
    Rem L'utilisateur a cliqué sur Annuler.
    if sRecipient = "" then
        exit sub
    endif

    Rem Demandez à l'utilisateur le sujet du courriel.
    sSubject = InputBox("Veuillez saisir l'objet du courriel")
    Rem L'utilisateur a cliqué sur Annuler.
    if sSubject = "" then
        exit sub
    endif

    Rem Demander à l'utilisateur le contenu du courriel.
    sBody = InputBox("Veuillez saisir le contenu du courriel")
    Rem L'utilisateur a cliqué sur Annuler.
    if sBody = "" then
        exit sub
    endif

    Rem Ok, maintenant que nous avons tout, nous commençons à envoyer un email.

    Rem Nous utiliserons 4 services UNO qui sont:
    Rem - com.sun.star.mail.MailUser: C'est le service qui va assurer la bonne configuration
    Rem des serveurs SMTP et IMAP (on peut remercier Mozilla pour la base de données ISPBD que j'utilise).
    Rem - com.sun.star.mail.MailServiceProvider: il s'agit du service qui vous permet d'utiliser les serveurs
    Rem SMTP et IMAP. Nous utiliserons ce service à l'aide du service précédent.
    Rem - com.sun.star.datatransfer.TransferableFactory: Ce service est une forge pour la création de
    Rem Transferable qui sont la base du corps de l'email ainsi que de ses fichiers joints.
    Rem - com.sun.star.mail.MailMessage: il s'agit du service qui implémente le message électronique.
    Rem Maintenant que tout est clair, nous pouvons commencer.


    Rem Nous créons d’abord l’email.

    Rem Il s'agit de notre forge de transférable, elle simplifie grandement l'API de messagerie de LibreOffice...
    oTransferable = createUnoService("com.sun.star.datatransfer.TransferableFactory")

    Rem oBody est le corps du courriel. Il est créé ici à partir d'une chaîne de caractères mais pourrait également
    Rem avoir été créé à partir d'un InputStream, d'une URL de fichier (file://...) ou d'une séquence d'octets.
    oBody = oTransferable.getByString(sBody)

    Rem oMail est le message électronique. Il est créé à partir du service com.sun.star.mail.MailMessage.
    Rem Il peut être créé avec une pièce jointe avec la méthode createWithAttachment().
    oMail = com.sun.star.mail.MailMessage.create(sRecipient, sSender, sSubject, oBody)

    Rem Demandez à l'utilisateur les URL des fichiers joints.
    oDialog = createUnoService("com.sun.star.ui.dialogs.FilePicker")
    oDialog.setMultiSelectionMode(true)
    if oDialog.execute() = com.sun.star.ui.dialogs.ExecutableDialogResults.OK then
        oFiles() = oDialog.getSelectedFiles()
        Rem Ces deux services sont simplement utilisés pour obtenir un nom de fichier approprié.
        oUrlTransformer = createUnoService("com.sun.star.util.URLTransformer")
        oUriFactory = createUnoService("com.sun.star.uri.UriReferenceFactory")
        for i = lbound(oFiles()) To ubound(oFiles())
            oUri = getUri(oUrlTransformer, oUriFactory, oFiles(i))
            oAttachment = createUnoStruct("com.sun.star.mail.MailAttachment")
            Rem Il faut saisir ReadableName. Il s'agit du nom du fichier joint
            Rem tel qu'il apparaît dans l'e-mail. Ici, nous obtenons le nom du fichier.
            oAttachment.ReadableName = oUri.getPathSegment(oUri.getPathSegmentCount() - 1)
            Rem La pièce jointe est récupérée à partir d'une URL mais comme pour oBody elle peut être
            Rem récupérée à partir d'une chaîne de caractères, d'un InputStream ou d'une séquence d'octets.
            oAttachment.Data = oTransferable.getByUrl(oUri.getUriReference())
            oMail.addAttachment(oAttachment)
            next i
    endif
    Rem Fin de la création du courriel.


    Rem Maintenant, nous devons envoyer le courriel.

    Rem Nous créons d'abord un MailUser à partir de l'adresse de l'expéditeur. Il ne s'agit pas nécessairement de
    Rem l'adresse de l'expéditeur, mais elle doit suivre la rfc822 (ie: surnom <nom@fai.com>).
    Rem L'assistant IspDB sera automatiquement lancé si cet utilisateur n'a jamais été configuré.
    oUser = com.sun.star.mail.MailUser.create(sSender)
    Rem L'utilisateur a annulé l'assistant IspDB.
    if isNull(oUser) then
        exit sub
    endif

    Rem Maintenant que nous avons l’utilisateur, nous pouvons vérifier s’il souhaite utiliser une adresse de réponse.
    if oUser.useReplyTo() then
        oMail.ReplyToAddress = oUser.getReplyToAddress()
    endif
    Rem De la même manière, je peux tester si l'utilisateur a une configuration IMAP avec oUser.supportIMAP()
    Rem puis créer un courriel de regroupement si nécessaire. Dans ce cas, vous devez :
    Rem - Construire un courriel de regroupement (comme précédemment pour oMail).
    Rem - Créer et vous connecter à un serveur IMAP (comme nous le ferons pour SMTP).
    Rem - Télécharger ce courriel sur le serveur IMAP avec: oServer.uploadMessage(oServer.getSentFolder(), oMail).
    Rem - Une fois téléchargé, récupérer le MessageId avec la propriété oMail.MessageId.
    Rem - Définir la propriété oMail.ThreadId sur MessageId pour tous les courriels suivants.
    Rem Super, vous avez réussi à regrouper l'envoi de courriels dans un courriel de regroupement.

    Rem Pour envoyer l'e-mail, nous devons créer un serveur SMTP. Voici comment procéder :
    SMTP = com.sun.star.mail.MailServiceType.SMTP
    oServer = createUnoService("com.sun.star.mail.MailServiceProvider").create(SMTP)
    Rem Nous nous connectons maintenant en utilisant la configuration SMTP de l'utilisateur.
    oServer.connect(oUser.getConnectionContext(SMTP), oUser.getAuthenticator(SMTP))
    Rem Et bien ça y est, nous sommes connectés, il ne reste plus qu'à envoyer le courriel.
    oServer.sendMailMessage(oMail)
    Rem N'oubliez pas de fermer la connexion.
    oServer.disconnect()
    MsgBox "Le courriel a été envoyé avec succès." & chr(13) & "Son MessageId est: " & oMail.MessageId
    Rem Et voilà, le tour est joué...

End Sub


Function getUri(oUrlTransformer As Variant, oUriFactory As Variant, sUrl As String) As Variant
    oUrl = createUnoStruct("com.sun.star.util.URL")
    oUrl.Complete = sUrl
    oUrlTransformer.parseStrict(oUrl)
    oUri = oUriFactory.parse(oUrlTransformer.getPresentation(oUrl, false))
    getUri = oUri
End Function
```

Et voilà, le tour est joué, cela n'a pris que quelques lignes de code alors profitez-en...  
Par contre, ce n'est qu'un exemple de vulgarisation, et tous les contrôles d'erreurs nécessaires ne sont pas en place...

___

## Comment personnaliser les menus de LibreOffice:

Afin de vous permettre de placer l'accès aux différentes fonctionnalités d'eMailerOOo où vous le souhaitez, il est désormais possible de créer des menus personnalisés pour les commandes:
- `ShowIspdb` pour **Configurer la connexion** avec l'Étendue LibreOffice.
- `ShowMailer` pour **Envoyer un document** avec l'Étendue LibreOffice.
- `ShowMerger` pour **Publiposter un document** avec l'Étendue Writer et Calc.
- `ShowSpooler` pour **Courriels sortants** avec l'Étendue LibreOffice.
- `StartSpooler` pour **Démarrer le spouleur** avec l'Étendue LibreOffice.
- `StopSpooler` pour **Arrêter le spouleur** avec l'Étendue LibreOffice.

Dans l'onglet **Menu** de la fenêtre **Outils -> Personnaliser**, sélectionnez **Macros** dans **Catégorie** pour accéder aux macros sous: **Mes macros -> eMailerOOo**.  
Vous devrez peut-être ouvrir les applications (Writer et Calc) et ajouter les macros avec l'**Étendue** définie sur les applications prises en charge.

Cela ne doit être fait qu'une seule fois pour LibreOffice et chaque application, et malheureusement je n'ai encore rien trouvé de plus simple.

___

## Comment créer l'extension:

Normalement, l'extension est créée avec Eclipse pour Java et [LOEclipse][74]. Pour contourner Eclipse, j'ai modifié LOEclipse afin de permettre la création de l'extension avec Apache Ant.  
Pour créer l'extension eMailerOOo avec l'aide d'Apache Ant, vous devez:
- Installer le [SDK Java][75] version 8 ou supérieure.
- Installer [Apache Ant][76] version 1.10.0 ou supérieure.
- Installer [LibreOffice et son SDK][77] version 7.x ou supérieure.
- Cloner le dépôt [eMailerOOo][78] sur GitHub dans un dossier.
- Depuis ce dossier, accédez au répertoire: `source/eMailerOOo/`
- Dans ce répertoire, modifiez le fichier `build.properties` afin que les propriétés `office.install.dir` et `sdk.dir` pointent vers les dossiers d'installation de LibreOffice et de son SDK, respectivement.
- Lancez la création de l'archive avec la commande: `ant`
- Vous trouverez l'archive générée dans le sous-dossier: `dist/`

___

## A été testé avec:

* LibreOffice 7.3.7.2 - Lubuntu 22.04 - Python version 3.10.12 - OpenJDK-11-JRE (amd64)

* LibreOffice 7.5.4.2(x86) - Windows 10 - Python version 3.8.16 - Adoptium JDK Hotspot 11.0.19 (sous Lubuntu 22.04 / VirtualBox 6.1.38)

* LibreOffice 7.4.3.2(x64) - Windows 10(x64) - Python version 3.8.15  - Adoptium JDK Hotspot 11.0.17 (x64) (sous Lubuntu 22.04 / VirtualBox 6.1.38)

* LibreOffice 24.2.1.2 - Lubuntu 22.04

* LibreOffice 24.8.0.3 (X86_64) - Windows 10(x64) - Python version 3.9.19 (sous Lubuntu 22.04 / VirtualBox 6.1.38)

* **Ne fonctionne pas avec OpenOffice** voir [dysfonctionnement 128569][79]. N'ayant aucune solution, je vous encourrage d'installer **LibreOffice**.

Je vous encourage en cas de problème :confused:  
de créer un [dysfonctionnement][12]  
J'essaierai de le résoudre :smile:

___

## Historique:

### [Toutes les changements sont consignées dans l'Historique des versions][80]

[1]: </img/emailer.svg#collapse>
[2]: <https://prrvchr.github.io/eMailerOOo/>
[3]: <https://prrvchr.github.io/eMailerOOo/>
[4]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/TermsOfUse_fr>
[5]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/PrivacyPolicy_fr>
[6]: <https://prrvchr.github.io/eMailerOOo/README_fr#ce-qui-a-%C3%A9t%C3%A9-fait-pour-la-version-152>
[7]: <https://prrvchr.github.io/>
[8]: <https://www.libreoffice.org/download/download-libreoffice/>
[9]: <https://www.openoffice.org/download/index.html>
[10]: <https://prrvchr.github.io/eMailerOOo/README_fr#envoi-de-courriel-avec-une-macro-libreoffice-en-basic>
[11]: <https://github.com/prrvchr/eMailerOOo>
[12]: <https://github.com/prrvchr/eMailerOOo/issues/new>
[13]: <https://github.com/sponsors/prrvchr>
[14]: <https://appdefensealliance.dev/casa>
[15]: <https://github.com/prrvchr/OAuth2OOo/blob/master/LOV_OAuth2OOo.pdf>
[16]: <https://img.shields.io/static/v1?label=Sponsor&message=%E2%9D%A4&logo=GitHub&color=%23fe8e86#right>
[17]: <https://prrvchr.github.io/OAuth2OOo/README_fr#pr%C3%A9requis>
[18]: <https://prrvchr.github.io/jdbcDriverOOo/README_fr#pr%C3%A9requis>
[19]: <https://prrvchr.github.io/OAuth2OOo/img/OAuth2OOo.svg#middle>
[20]: <https://prrvchr.github.io/OAuth2OOo/README_fr>
[21]: <https://github.com/prrvchr/OAuth2OOo/releases/latest/download/OAuth2OOo.oxt>
[22]: <https://img.shields.io/github/v/tag/prrvchr/OAuth2OOo?label=latest#right>
[23]: <https://prrvchr.github.io/jdbcDriverOOo/img/jdbcDriverOOo.svg#middle>
[24]: <https://prrvchr.github.io/jdbcDriverOOo/README_fr>
[25]: <https://github.com/prrvchr/jdbcDriverOOo/releases/latest/download/jdbcDriverOOo.oxt>
[26]: <https://img.shields.io/github/v/tag/prrvchr/jdbcDriverOOo?label=latest#right>
[27]: <https://prrvchr.github.io/vCardOOo/img/vCardOOo.svg#middle>
[28]: <https://prrvchr.github.io/vCardOOo/README_fr>
[29]: <https://github.com/prrvchr/vCardOOo/releases/latest/download/vCardOOo.oxt>
[30]: <https://img.shields.io/github/v/tag/prrvchr/vCardOOo?label=latest#right>
[31]: <https://fr.wikipedia.org/wiki/Nextcloud>
[32]: <https://prrvchr.github.io/gContactOOo/img/gContactOOo.svg#middle>
[33]: <https://prrvchr.github.io/gContactOOo/README_fr>
[34]: <https://github.com/prrvchr/gContactOOo/releases/latest/download/gContactOOo.oxt>
[35]: <https://img.shields.io/github/v/tag/prrvchr/gContactOOo?label=latest#right>
[36]: <https://prrvchr.github.io/mContactOOo/img/mContactOOo.svg#middle>
[37]: <https://prrvchr.github.io/mContactOOo/README_fr>
[38]: <https://github.com/prrvchr/mContactOOo/releases/latest/download/mContactOOo.oxt>
[39]: <https://img.shields.io/github/v/tag/prrvchr/mContactOOo?label=latest#right>
[40]: <https://prrvchr.github.io/HyperSQLOOo/img/HyperSQLOOo.svg#middle>
[41]: <https://prrvchr.github.io/HyperSQLOOo/README_fr>
[42]: <https://github.com/prrvchr/HyperSQLOOo/releases/latest/download/HyperSQLOOo.oxt>
[43]: <https://img.shields.io/github/v/tag/prrvchr/HyperSQLOOo?label=latest#right>
[44]: <https://prrvchr.github.io/HyperSQLOOo/README_fr#comment-importer-des-donn%C3%A9es-depuis-un-fichier-calc>
[45]: <https://prrvchr.github.io/eMailerOOo/img/eMailerOOo.svg#middle>
[46]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/eMailerOOo.oxt>
[47]: <https://img.shields.io/github/downloads/prrvchr/eMailerOOo/latest/total?label=v1.5.1#right>
[48]: <../setup/fr/>
[49]: <https://prrvchr.github.io/eMailerOOo/README_fr#publipostage-de-courriels-avec-des-listes-de-diffusion>
[50]: <https://prrvchr.github.io/eMailerOOo/README_fr#configuration-de-la-connexion>
[51]: <https://prrvchr.github.io/eMailerOOo/README_fr#courriels-sortants>
[52]: <img/eMailerOOo-Merger1_fr.png>
[53]: <img/eMailerOOo-Merger2_fr.png>
[54]: <img/eMailerOOo-Merger3_fr.png>
[55]: <img/eMailerOOo-Merger4_fr.png>
[56]: <https://prrvchr.github.io/eMailerOOo/README_fr#destinataires-disponibles>
[57]: <img/eMailerOOo-Merger5_fr.png>
[58]: <img/eMailerOOo-Merger6_fr.png>
[59]: <img/eMailerOOo-Merger7_fr.png>
[60]: <img/eMailerOOo-Merger8_fr.png>
[61]: <img/eMailerOOo-Merger9_fr.png>
[62]: <img/eMailerOOo-Merger10_fr.png>
[63]: <img/eMailerOOo-Merger11_fr.png>
[64]: <img/eMailerOOo-Ispdb1_fr.png>
[65]: <img/eMailerOOo-Ispdb2_fr.png>
[66]: <img/eMailerOOo-Ispdb3_fr.png>
[67]: <img/eMailerOOo-Ispdb4_fr.png>
[68]: <img/eMailerOOo-Ispdb5_fr.png>
[69]: <img/eMailerOOo-Ispdb6_fr.png>
[70]: <img/eMailerOOo-Spooler1_fr.png>
[71]: <https://prrvchr.github.io/eMailerOOo/README_fr#journal-dactivité-du-spouleur>
[72]: <img/eMailerOOo-Spooler2_fr.png>
[73]: <img/eMailerOOo-Spooler3_fr.png>
[74]: <https://github.com/LibreOffice/loeclipse>
[75]: <https://adoptium.net/temurin/releases/?version=8&package=jdk>
[76]: <https://ant.apache.org/manual/install.html>
[77]: <https://downloadarchive.documentfoundation.org/libreoffice/old/7.6.7.2/>
[78]: <https://github.com/prrvchr/eMailerOOo.git>
[79]: <https://bz.apache.org/ooo/show_bug.cgi?id=128569>
[80]: <../change/fr/>
