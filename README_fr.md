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

# version [1.4.3][6]

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

Bref, à participer au developpement de cette extension.  
Car c'est ensemble que nous pouvons rendre le Logiciel Libre plus intelligent.

___

## Prérequis:

L'extension eMailerOOo utilise l'extension OAuth2OOo pour fonctionner.  
Elle doit donc répondre aux [prérequis de l'extension OAuth2OOo][13].

L'extension eMailerOOo utilise l'extension jdbcDriverOOo pour fonctionner.  
Elle doit donc répondre aux [prérequis de l'extension jdbcDriverOOo][14].  
De plus, eMailerOOo nécessite que l'extension jdbcDriverOOo soit configurée pour fournir `com.sun.star.sdb` comme niveau d'API, qui est la configuration par défaut.

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

        Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts présents sur une plateforme [Nextcloud][29] comme source de données pour les listes de diffusion et la fusion de documents.

    - [![gContactOOo logo][30]][31] Installer l'extension **[gContactOOo.oxt][32]** [![Version][33]][32]

        Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts téléphoniques personnels (contact Android) comme source de données pour les listes de diffusion et la fusion de documents.

    - [![mContactOOo logo][34]][35] Installer l'extension **[mContactOOo.oxt][36]** [![Version][37]][36]

        Cette extension n'est nécessaire que si vous souhaitez utiliser vos contacts Microsoft Outlook comme source de données pour les listes de diffusion et la fusion de documents.

    - [![HyperSQLOOo logo][38]][39] Installer l'extension **[HyperSQLOOo.oxt][40]** [![Version][41]][40]

        Cette extension n'est nécessaire que si vous souhaitez utiliser un fichier Calc comme source de données pour les listes de diffusion et la fusion de documents. Voir: [Comment importer des données depuis un fichier Calc][42].

- ![eMailerOOo logo][43] Installer l'extension **[eMailerOOo.oxt][44]** [![Version][45]][44]

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
- [vCardOOo][26]. Cette extension vous permettra d'utiliser vos contacts présents sur une plateforme [Nextcloud][29] comme source de données.
- [gContactOOo][31]. Cette extension vous permettra d'utiliser votre téléphone Android (vos contacts téléphoniques) comme source de données.
- [mContactOOo][35]. Cette extension vous permettra d'utiliser vos contacts Microsoft Outlook comme source de données.
- [HyperSQLOOo][39]. Cette extension vous permettra d'utiliser un fichier Calc comme source de données. Voir: [Comment importer des données depuis un fichier Calc][42].

Pour les 3 premières extensions le nom de la **table principale** peut être trouvé (et même changé avant toute connexion) dans:  
**Outils -> Options -> Internet -> Nom de l'extension -> Nom de la table principale**

Ce mode d'utilisation est composé de 3 sections:
- [Publipostage de courriels avec des listes de diffusion][46].
- [Configuration de la connexion][47].
- [Courriels sortants][48].

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

![eMailerOOo Merger screenshot 1][49]

#### Sélection de la source de données:

Le chargement de la source de données de l'assistant **Publipostage de courriels** devrait apparaître :

![eMailerOOo Merger screenshot 2][50]

Les captures d'écran suivantes utilisent l'extension [gContactOOo][31] comme **source de données**. Si vous utilisez votre propre **source de données**, il est nécessaire d'adapter les paramètres par rapport à celle-ci. 

Dans la copie d'écran suivante, on peut voir que la **source de données** gContactOOo s'appelle: `Adresses` et que dans la liste des tables la table: `PUBLIC.Tous mes contacts` est sélectionnée.

![eMailerOOo Merger screenshot 3][51]

Si aucune liste de diffusion n'existe, vous devez en créer une, en saisissant son nom et en validant avec: `ENTRÉE` ou le bouton `Ajouter`.

Assurez-vous lors de la création de la liste de diffusion que la **table principale** est toujours bien sélectionnée dans la liste des tables.  
Si cette recommandation n'est pas suivie alors **la fusion de documents ne fonctionnera pas** et ceci silencieusement.

![eMailerOOo Merger screenshot 4][52]

Maintenant que votre nouvelle liste de diffusion est disponible dans la liste, vous devez la sélectionner.

Et ajouter les colonnes suivantes:
- Colonne de clef primaire: `Uri`
- Colonnes d'adresses électronique: `HomeEmail`, `WorkEmail` et `OtherEmail`

Si plusieurs colonnes d'adresses courriel sont sélectionnées, alors l'ordre devient pertinent puisque le courriel sera envoyé à la première adresse disponible.  
De plus, à l'étape Sélection des destinataires de l'assistant, dans l'onglet [Destinataires disponibles][53], seuls les enregistrements avec au moins une colonne d'adresse courriel saisie seront répertoriés.  
Assurez-vous donc d'avoir un carnet d'adresses avec au moins un des champs d'adresse e-mail (Home, Work ou Other) renseigné.

![eMailerOOo Merger screenshot 5][54]

Ce paramètrage ne doit être effectué que pour les nouvelles listes de diffusion.  
Vous pouvez maintenant passer à l'étape suivante.

#### Sélection des destinataires:

##### Destinataires disponibles:

Les destinataires sont sélectionnés à l'aide de 2 boutons `Tout ajouter` et `Ajouter` permettant respectivement:
- Soit d'ajouter le groupe de destinataires sélectionnés dans la liste `Carnet d'adresses`. Ceci permet lors d'un publipostage, que les modifications du contenu du groupe soient prises en compte. Une liste de diffusion n'accepte qu'un seul groupe.
- Soit d'ajouter la sélection, qui peut être multiple à l'aide de la touche `CTRL`. Cette sélection est immuable quelle que soit la modification des groupes du carnet d'adresses.

![eMailerOOo Merger screenshot 6][55]

Example de la sélection multiple:

![eMailerOOo Merger screenshot 7][56]

##### Destinataires sélectionnés:

Les destinataires sont désélectionnés à l'aide de 2 boutons `Tout retirer` et `Retirer` permettant respectivement:
- Soit de retirer le groupe qui a été affecté à cette liste de diffusion. Ceci est nécessaire afin de pouvoir modifier à nouveau le contenu de cette liste de diffusion.
- Soit de retirer la sélection, qui peut être multiple à l'aide de la touche `CTRL`.

![eMailerOOo Merger screenshot 8][57]

Si vous avez sélectionné au moins 1 destinataire, vous pouvez passer à l'étape suivante.

#### Sélection des options d'envoi:

Si cela n'est pas déjà fait, vous devez créer un nouvel expéditeur à l'aide du bouton `Ajouter`.

![eMailerOOo Merger screenshot 9][58]

La création du nouvel expéditeur est décrite dans la section [Configuration de la connexion][47].

Le courriel doit avoir un sujet. Il peut être enregistré dans le document Writer.  
Vous pouvez insérer des champs de fusion dans l'objet du courriel. Un champ de fusion est composé d'une accolade ouvrante, du nom de la colonne référencée (sensible à la casse) et d'une accolade fermante (ie: `{NomColonne}`).

![eMailerOOo Merger screenshot 10][59]

Le courriel peut éventuellement contenir des fichiers joints. Ils peuvent être enregistrés dans le document Writer.  
La capture d'écran suivante montre 1 fichier joint qui sera fusionné sur la source de données puis converti au format PDF avant d'être joint au courriel.

![eMailerOOo Merger screenshot 11][60]

Assurez-vous de toujours quitter l'assistant avec le bouton `Terminer` pour confirmer la soumission des travaux d'envoi.  
Pour envoyer les travaux d'envoi, veuillez suivre la section [Courriels sortants][48].

### Configuration de la connexion:

#### Démarrage de l'assistant de connexion:

Dans LibreOffice aller à: **Outils -> Add-ons -> Envoi de courriels -> Configurer la connexion**

![eMailerOOo Ispdb screenshot 1][61]

#### Sélection du compte:

![eMailerOOo Ispdb screenshot 2][62]

#### Trouver la configuration:

![eMailerOOo Ispdb screenshot 3][63]

#### Configuration SMTP:

![eMailerOOo Ispdb screenshot 4][64]

#### Configuration IMAP:

![eMailerOOo Ispdb screenshot 5][65]

#### Tester la connexion:

![eMailerOOo Ispdb screenshot 6][66]

Assurez-vous de toujours quitter l'assistant avec le bouton `Terminer` afin d'enregistrer les paramètres de connexion.

### Courriels sortants:

#### Démarrage du spouleur de courriels:

Dans LibreOffice aller à: **Outils -> Add-ons -> Envoi de courriels -> Courriels sortants**

![eMailerOOo Spooler screenshot 1][67]

#### Liste des courriels sortants:

Chaque travaux d'envoi possède 3 états différents:
- État **0**: le courriel est prêt à être envoyé.
- État **1**: le courriel a été envoyé avec succès.
- État **2**: Une erreur est survenue lors de l'envoi du courriel. Vous pouvez consulter le message d'erreur dans le [Journal d'activité du spouleur][68].

![eMailerOOo Spooler screenshot 2][69]

Le spouleur de courriels est arrêté par défaut. **Il doit être démarré avec le bouton `Démarrer / Arrêter` pour que les courriels en attente soient envoyés**.

#### Journal d'activité du spouleur:

Lorsque le spouleur de courriel est démarré, son activité peut être visualisée dans le journal d'activité.

![eMailerOOo Spooler screenshot 3][70]

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

Normalement, l'extension est créée avec Eclipse pour Java et [LOEclipse][28]. Pour contourner Eclipse, j'ai modifié LOEclipse afin de permettre la création de l'extension avec Apache Ant.  
Pour créer l'extension eMailerOOo avec l'aide d'Apache Ant, vous devez:
- Installer le [SDK Java][29] version 8 ou supérieure.
- Installer [Apache Ant][30] version 1.10.0 ou supérieure.
- Installer [LibreOffice et son SDK][31] version 7.x ou supérieure.
- Cloner le dépôt [eMailerOOo][32] sur GitHub dans un dossier.
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

* **Ne fonctionne pas avec OpenOffice** voir [dysfonctionnement 128569][71]. N'ayant aucune solution, je vous encourrage d'installer **LibreOffice**.

Je vous encourage en cas de problème :confused:  
de créer un [dysfonctionnement][12]  
J'essaierai de le résoudre :smile:

___

## Historique:

### Ce qui a été fait pour la version 0.0.1:

- Ecriture de [IspDB][72] ou l'assistant de configuration de connexion aux serveurs SMTP permettant:
    - De trouver les paramètres de connexion à un serveur SMTP à partir d'une adresse courriel. D'ailleur je remercie particulierement Mozilla, pour [Thunderbird autoconfiguration database][73] ou IspDB, qui à rendu ce défi possible...
    - D'afficher l'activité du service UNO `com.sun.star.mail.MailServiceProvider` lors de la connexion au serveur SMTP et l'envoi d'un courriel. 

- Ecriture du [Spouleur][74] de courriels permettant:
    - D'afficher les travaux d'envoi de courriel avec leurs états respectifs.
    - D'afficher l'activité du service UNO `com.sun.star.mail.SpoolerService` lors de l'envoi de courriels.
    - De démarrer et arrêter le service spouleur.

- Ecriture du [Merger][75] ou l'assistant de publipostage de courriels permettant:
    - De créer des listes de diffusions.
    - De fusionner et convertir au format HTML le document courant pour en faire le message du courriel.
    - De fusionner et/ou convertir au format PDF d'éventuel fichiers joints au courriel.

- Ecriture du [Mailer][76] de document permettant:
    - De convertir au format HTML le document pour en faire le message du courriel.
    - De convertir au format PDF d'éventuel fichiers joints au courriel.

- Ecriture d'un [Grid][77] piloté par un `com.sun.star.sdb.RowSet` permettant:
    - D'être paramètrable sur les colonnes à afficher.
    - D'être paramètrable sur l'ordre de tri à afficher.
    - De sauvegarder les paramètres d'affichage.

### Ce qui a été fait pour la version 0.0.2:

- Réécriture de [IspDB][72] ou Assistant de configuration de connexion aux serveurs de messagerie afin d'intégrer la configuration de la connexion IMAP.
    - Utilisation de [IMAPClient][78] version 2.2.0: une bibliothèque cliente IMAP complète, Pythonic et facile à utiliser.
    - Extension des fichiers IDL [com.sun.star.mail.*][79]:
        - [XMailMessage2.idl][80] prend désormais en charge la hiérarchisation des courriels (thread).
        - La nouvelle interface [XImapService][81] permet d'accéder à une partie de la bibliothèque IMAPClient.

- Réécriture du [Spouleur][82] afin d'intégrer des fonctionnalités IMAP comme la création d'un fil récapitulant le publipostage et regroupant tous les courriels envoyés.

- Soumission de l'extension eMailerOOo à Google et obtention de l'autorisation d'utiliser son API GMail afin d'envoyer des courriels avec un compte Google.

### Ce qui a été fait pour la version 0.0.3:

- Réécriture du [Grid][77] afin de permettre:
    - Le tri sur une colonne avec l'intégration du service UNO [SortableGridDataModel][83].
    - La génération des filtres des enregistrements nécessaires au service [Spouleur][74].
    - Le partage avec le module python [Grid][84] de l'extension [jdbcDriverOOo][22].

- Réécriture du [Merger][75] afin de permettre:
    - La gestion du nom du Schema dans de nom des tables afin d'être compatible avec la version 0.0.4 de [jdbcDriverOOo][22].
    - La création de liste de diffusion sur un groupe du carnet d'adresse et permettant de suivre la modification de son contenu.
    - L'utilisation de clé primaire, qui peuvent être composite, supportant les [DataType][85] `VARCHAR` et `INTEGER` ou derivé.
    - Un aperçu du document avec des champs de fusion remplis plus rapidement grâce au [Grid][77].

- Réécriture du [Spouleur][74] afin de permettre:
    - L'utilisation des nouveaux filtres supportant les clés primaires composite fourni par le [Merger][75].
    - L'utilisation du nouveau [Grid][77] permettant le tri sur une colonne.

- Encore plein d'autres choses...

### Ce qui a été fait pour la version 1.0.0:

- L'extension **smtpMailerOOo** a été renomé en **eMailerOOo**.

### Ce qui a été fait pour la version 1.0.1:

- L'absence ou l'obsolescence des extensions **OAuth2OOo** et/ou **jdbcDriverOOo** nécessaires au bon fonctionnement de **eMailerOOo** affiche désormais un message d'erreur. Ceci afin d'éviter qu'un dysfonctionnement tel que le [dysfonctionnement #3][86] ne se reproduise...

- La base de données HsqlDB sous-jacente peut être ouverte dans Base avec: **Outils -> Options -> Internet -> eMailerOOo -> Base de données**.

- Le menu **Outils -> Add-ons** s'affiche désormais correctement en fonction du contexte.

- Encore plein d'autres choses...

### Ce qui a été fait pour la version 1.0.2:

- Si aucune configuration n'est trouvée dans l'assistant de configuration de la connexion (IspDB Wizard) alors il est possible de configurer la connexion manuellement. Voir [dysfonctionnement #5][87].

### Ce qui a été fait pour la version 1.1.0:

- Dans l'assistant de configuration de la connexion (IspDB Wizard) il est maintenant possible de désactiver la configuration IMAP.  
    En conséquence, cela n'envoie plus de fil de discussion (message IMAP) lors de la fusion d'un mailing.  
    Dans ce même assistant, il est désormais possible de saisir une adresse courriel de réponse.

- Dans l'assistant de fusion d'email, il est désormais possible d'insérer des champs de fusion dans l'objet du courriel. Voir [dysfonctionnement #6][88].  
    Dans le sujet d'un courriel, un champ de fusion est composé d'une accolade ouvrante, du nom de la colonne référencée (sensible à la casse) et d'une accolade fermante (ie: `{NomDeLaColonne}`).  
    Lors de la saisie du sujet du courriel, une erreur de syntaxe dans un champ de fusion sera signalée et empêchera la soumission du mailing.

- Il est désormais possible dans le Spouleur de visualiser les courriels au format eml.

- Un service [com.sun.star.mail.MailUser][89] permet désormais d'accéder à une configuration de connexion (SMTP et/ou IMAP) depuis une adresse courriel qui suite la rfc822.  
    Un autre service [com.sun.star.datatransfer.TransferableFactory][90] permet, comme son nom l'indique, la création de [Transferable][91] à partir d'un texte (string), d'une séquence binaire, d'une Url (file://...) ou un flux de données (InputStream).  
    Ces deux nouveaux services simplifient grandement l'API mail de LibreOffice et permettent d'envoyer des courriels depuis Basic. Voir le [dysfonctionnement #4][92].  
    Vous trouverez une macro Basic vous permettant d'envoyer des emails dans : **Outils -> Macros -> Editer les Macros... -> eMailerOOo -> SendEmail**.

### Ce qui a été fait pour la version 1.1.1:

- Prise en charge de la version 1.2.0 de l'extension **OAuth2OOo**. Les versions précédentes ne fonctionneront pas avec l'extension **OAuth2OOo** 1.2.0 ou ultérieure.

### Ce qui a été fait pour la version 1.2.0:

- Tous les paquets Python nécessaires à l'extension sont désormais enregistrés dans un fichier [requirements.txt][93] suivant la [PEP 508][94].
- Désormais si vous n'êtes pas sous Windows alors les paquets Python nécessaires à l'extension peuvent être facilement installés avec la commande:  
  `pip install requirements.txt`
- Modification de la section [Prérequis][95].

### Ce qui a été fait pour la version 1.2.1:

- Correction d'une régression permettant l'affichage des erreurs dans le Spouleur.
- Intégration d'un correctif pour contourner le [dysfonctionnement #159988][96].

### Ce qui a été fait pour la version 1.2.2:

- La création de la base de données, lors de la première connexion, utilise l'API UNO proposée par l'extension jdbcDriverOOo depuis la version 1.3.2. Cela permet d'enregistrer toutes les informations nécessaires à la création de la base de données dans 5 tables texte qui sont en fait [5 fichiers csv][97].
- L'extension vous demandera d'installer les extensions OAuth2OOo et jdbcDriverOOo en version respectivement 1.3.4 et 1.3.2 minimum.
- De nombreuses corrections.

### Ce qui a été fait pour la version 1.2.3:

- Correction d'une régression provenant de la version 1.2.2 et empêchant la soumission des travaux dans le spouleur de courriels.
- Correction du [dysfonctionnement #7][98] ne permettant pas l'affichage des messages d'erreur en cas de configuration incorrecte.

### Ce qui a été fait pour la version 1.2.4:

- Mise à jour du paquet [Python decorator][99] vers la version 5.1.1.
- Mise à jour du paquet [Python ijson][100] vers la version 3.3.0.
- Mise à jour du paquet [Python packaging][101] vers la version 24.1.
- Mise à jour du paquet [Python setuptools][102] vers la version 72.1.0 afin de répondre à l'[alerte de sécurité Dependabot][103].
- Mise à jour du paquet [Python validators][104] vers la version 0.33.0.
- L'extension vous demandera d'installer les extensions OAuth2OOo et jdbcDriverOOo en version respectivement 1.3.6 et 1.4.2 minimum.

### Ce qui a été fait pour la version 1.2.5:

- Mise à jour du paquet [Python setuptools][102] vers la version 73.0.1.
- L'extension vous demandera d'installer les extensions OAuth2OOo et jdbcDriverOOo en version respectivement 1.3.7 et 1.4.5 minimum.
- Les modifications apportées aux options de l'extension, qui nécessitent un redémarrage de LibreOffice, entraîneront l'affichage d'un message.
- Support de LibreOffice version 24.8.x.

### Ce qui a été fait pour la version 1.2.6:

- Si une adresse de réponse a été fournie, elle sera utilisée lors de la génération du fichier eml par le spouleur.
- L'extension vous demandera d'installer les extensions OAuth2OOo et jdbcDriverOOo en version respectivement 1.3.8 et 1.4.6 minimum.
- Modification des options de l'extension accessibles via : **Outils -> Options... -> Internet -> eMailerOOo** afin de respecter la nouvelle charte graphique.

### Ce qui a été fait pour la version 1.2.7:

- Le spouleur permet d'ouvrir les e-mails envoyés soit dans le client de messagerie local (ie: Thunderbird) soit en ligne dans votre navigateur pour les comptes utilisant une API d'envoi du courriel (ie: Google et Microsoft).
- Un nouvel onglet a été ajouté au spouleur pour permettre le suivi de l'activité du service de messagerie.
- Les connexions aux serveurs de messagerie Microsoft, qui ne fonctionnaient apparemment plus, ont été migrées vers l'API Graph.
- Pour les serveurs qui n'utilisent plus les protocoles SMTP et IMAP et proposent une API de remplacement (ie: Google API et Microsoft Graph):
    - Tous les paramètres des requêtes HTTP nécessaires à l'envoi de courriels sont stockés dans les fichiers de configuration de LibreOffice.
    - Toutes les données nécessaires au traitement des réponses HTTP sont stockées dans les fichiers de configuration de LibreOffice.

    Cela devrait permettre d'implémenter une API tierce pour l'envoi de courriels simplement en modifiant le fichier de configuration [Options.xcu][105].
- Pour fonctionner, ces nouvelles fonctionnalités nécessitent l'extension OAuth2OOo en version 1.3.9 minimum.
- La commande permettant d'ouvrir un courriel dans Thunderbird ne peut actuellement être modifiée que dans la configuration de LibreOffice (ie: Outils -> Options... -> Avancé -> Ouvrir la configuration avancée)
- Le non rafraîchissement des barres de défilement dans les listes multicolonnes (ie: grid) a été corrigé et sera disponible à partir de LibreOffice 24.8.4, voir [SortableGridDataModel cannot be notified for changes][106].
- L'ouverture des courriels dans votre navigateur ne fonctionne pas avec un compte Microsoft, l'url permettant cela n'a pas encore été trouvée et il semble que ce ne serait pas possible (ie: le popup doit être ouvert par la fenêtre Outlook)?
- De nombreuses corrections.

### Ce qui a été fait pour la version 1.3.0:

- L'extension vous demandera d'installer les extensions OAuth2OOo et jdbcDriverOOo en version respectivement 1.4.0 et 1.4.6 minimum.
- Seuls les fournisseurs disposant d'une API tierce ou d'une authentification OAuth2 et disposant d'une entrée dans la configuration de LibreOffice proposeront l'authentification OAuth2 par défaut dans l'assistant de configuration de connexion (ie: IspDB Wizard).
- Les fournisseurs de messagerie `yahoo.com` et `aol.com` ont été intégrés. Afin de faciliter la configuration, un lien vers la page permettant la création d'un mot de passe d'application a été ajouté à l'assistant de configuration de connexion. Si vous pensez que des liens vers d'autres fournisseurs manquent, veuillez ouvrir un dysfonctionnement afin que je puisse les rajouter.
- Mise à jour du paquet [Python IMAPClient][78] vers la version 3.0.1.
- Grâce aux améliorations apportées au [plugin Eclipse][107], il est désormais possible de créer le fichier de l'extension en utilisant la ligne de commande et l'outil de création d'archive [Apache Ant][108], voir le fichier [build.xml][109].
- L'extension refusera de s'installer sous OpenOffice quelle que soit la version ou LibreOffice autre que 7.x ou supérieur.
- Ajout des fichiers binaires nécessaires aux bibliothèques Python pour fonctionner sous Linux et LibreOffice 24.8 (ie: Python 3.9).
- De nombreuses corrections.

### Ce qui a été fait pour la version 1.3.1:

- Mise à jour du paquet [Python packaging][101] vers la version 24.2.
- Mise à jour du paquet [Python setuptools][102] vers la version 75.8.0.
- Mise à jour du paquet [Python six][110] vers la version 1.17.0.
- Mise à jour du paquet [Python validators][104] vers la version 0.34.0.
- Support de Python version 3.13.

### Ce qui a été fait pour la version 1.4.0:

- Mise à jour du paquet [Python packaging][60] vers la version 25.0.
- Rétrogradage du paquet [Python setuptools][61] vers la version 75.3.2, afin d'assurer la prise en charge de Python 3.8.
- Déploiement de l'enregistrement passif permettant une installation beaucoup plus rapide des extensions et de différencier les services UNO enregistrés de ceux fournis par une implémentation Java ou Python. Cet enregistrement passif est assuré par l'extension [LOEclipse][28] via les [PR#152][67] et [PR#157][68].
- Modification de [LOEclipse][28] pour prendre en charge le nouveau format de fichier `rdb` produit par l'utilitaire de compilation `unoidl-write`. Les fichiers `idl` ont été mis à jour pour prendre en charge les deux outils de compilation disponibles: idlc et unoidl-write.
- Il est désormais possible de créer le fichier oxt de l'extension eMailerOOo uniquement avec Apache Ant et une copie du dépôt GitHub. La section [Comment créer l'extension][69] a été ajoutée à la documentation.
- Implémentation de [PEP 570][70] dans la [journalisation][71] pour prendre en charge les arguments multiples uniques.
- Pour garantir la création correcte de la base de données eMailerOOo, il sera vérifié que l'extension jdbcDriverOOo a `com.sun.star.sdb` comme niveau d'API.
- Écriture de macros pour pouvoir placer des menus personnalisés où vous le souhaitez. Pour faciliter la création de ces menus personnalisés, la section [Comment personnaliser les menus de LibreOffice][72] a été ajoutée à la documentation.
- Nécessite l'extension **jdbcDriverOOo en version 1.5.0 minimum**.
- Nécessite l'extension **OAuth2OOo en version 1.5.0 minimum**.

### Ce qui a été fait pour la version 1.4.1:

- Dans l'assistant de connexion, si l'adresse courriel donnée n'est pas trouvée dans Mozilla IspDB ou si vous êtes hors ligne, les noms de serveur peuvent être de simples noms d'hôtes et les ports valides s'étendront jusqu'à 65535. Ceci afin de répondre à l'[issue#10][111].
- Les problèmes d'actualisation de la deuxième page de l'assistant de connexion ont été résolus par l'utilisation du service UNO `com.sun.star.awt.AsyncCallback`.
- Nécessite l'extension **jdbcDriverOOo en version 1.5.4 minimum**.
- Nécessite l'extension **OAuth2OOo en version 1.5.1 minimum**.

### Ce qui a été fait pour la version 1.4.2:

- Support de LibreOffice 25.2.x et 25.8.x sous Windows 64 bits.
- Nécessite l'extension **OAuth2OOo en version 1.5.2 minimum**.

### Ce qui a été fait pour la version 1.4.3:

- Toutes les méthodes exécutées en arrière-plan utilisent désormais le service UNO [com.sun.star.awt.AsyncCallback][112] pour le rappel.
- Nécessite l'extension **jdbcDriverOOo en version 1.5.7 minimum**.
- Nécessite l'extension **OAuth2OOo en version 1.5.3 minimum**.

### Que reste-t-il à faire pour la version 1.4.3:

- Ajouter de nouvelles langues pour l’internationalisation...

- Tout ce qui est bienvenu...

[1]: </img/emailer.svg#collapse>
[2]: <https://prrvchr.github.io/eMailerOOo/>
[3]: <https://prrvchr.github.io/eMailerOOo/>
[4]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/TermsOfUse_fr>
[5]: <https://prrvchr.github.io/eMailerOOo/source/eMailerOOo/registration/PrivacyPolicy_fr>
[6]: <https://prrvchr.github.io/eMailerOOo/README_fr#ce-qui-a-%C3%A9t%C3%A9-fait-pour-la-version-143>
[7]: <https://prrvchr.github.io/>
[8]: <https://www.libreoffice.org/download/download-libreoffice/>
[9]: <https://www.openoffice.org/download/index.html>
[10]: <https://prrvchr.github.io/eMailerOOo/README_fr#envoi-de-courriel-avec-une-macro-libreoffice-en-basic>
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
[38]: <https://prrvchr.github.io/HyperSQLOOo/img/HyperSQLOOo.svg#middle>
[39]: <https://prrvchr.github.io/HyperSQLOOo/README_fr>
[40]: <https://github.com/prrvchr/HyperSQLOOo/releases/latest/download/HyperSQLOOo.oxt>
[41]: <https://img.shields.io/github/v/tag/prrvchr/HyperSQLOOo?label=latest#right>
[42]: <https://prrvchr.github.io/HyperSQLOOo/README_fr#comment-importer-des-donn%C3%A9es-depuis-un-fichier-calc>
[43]: <https://prrvchr.github.io/eMailerOOo/img/eMailerOOo.svg#middle>
[44]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/eMailerOOo.oxt>
[45]: <https://img.shields.io/github/downloads/prrvchr/eMailerOOo/latest/total?label=v1.4.3#right>
[46]: <https://prrvchr.github.io/eMailerOOo/README_fr#publipostage-de-courriels-avec-des-listes-de-diffusion>
[47]: <https://prrvchr.github.io/eMailerOOo/README_fr#configuration-de-la-connexion>
[48]: <https://prrvchr.github.io/eMailerOOo/README_fr#courriels-sortants>
[49]: <img/eMailerOOo-Merger1_fr.png>
[50]: <img/eMailerOOo-Merger2_fr.png>
[51]: <img/eMailerOOo-Merger3_fr.png>
[52]: <img/eMailerOOo-Merger4_fr.png>
[53]: <https://prrvchr.github.io/eMailerOOo/README_fr#destinataires-disponibles>
[54]: <img/eMailerOOo-Merger5_fr.png>
[55]: <img/eMailerOOo-Merger6_fr.png>
[56]: <img/eMailerOOo-Merger7_fr.png>
[57]: <img/eMailerOOo-Merger8_fr.png>
[58]: <img/eMailerOOo-Merger9_fr.png>
[59]: <img/eMailerOOo-Merger10_fr.png>
[60]: <img/eMailerOOo-Merger11_fr.png>
[61]: <img/eMailerOOo-Ispdb1_fr.png>
[62]: <img/eMailerOOo-Ispdb2_fr.png>
[63]: <img/eMailerOOo-Ispdb3_fr.png>
[64]: <img/eMailerOOo-Ispdb4_fr.png>
[65]: <img/eMailerOOo-Ispdb5_fr.png>
[66]: <img/eMailerOOo-Ispdb6_fr.png>
[67]: <img/eMailerOOo-Spooler1_fr.png>
[68]: <https://prrvchr.github.io/eMailerOOo/README_fr#journal-dactivité-du-spouleur>
[69]: <img/eMailerOOo-Spooler2_fr.png>
[70]: <img/eMailerOOo-Spooler3_fr.png>
[71]: <https://bz.apache.org/ooo/show_bug.cgi?id=128569>
[72]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/ispdb>
[73]: <https://wiki.mozilla.org/Thunderbird:Autoconfiguration>
[74]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler>
[75]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/merger>
[76]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/mailer>
[77]: <https://github.com/prrvchr/eMailerOOo/tree/master/uno/lib/uno/grid>
[78]: <https://github.com/mjs/imapclient#readme>
[79]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/idl/com/sun/star/mail>
[80]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailMessage2.idl>
[81]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XImapService.idl>
[82]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/spooler.py>
[83]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/grid/SortableGridDataModel.html>
[84]: <https://github.com/prrvchr/jdbcDriverOOo/tree/master/source/jdbcDriverOOo/service/pythonpath/jdbcdriver/grid>
[85]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/sdbc/DataType.html>
[86]: <https://github.com/prrvchr/eMailerOOo/issues/3>
[87]: <https://github.com/prrvchr/eMailerOOo/issues/5>
[88]: <https://github.com/prrvchr/eMailerOOo/issues/6>
[89]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailUser.idl>
[90]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/datatransfer/XTransferableFactory.idl>
[91]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/datatransfer/XTransferable.html>
[92]: <https://github.com/prrvchr/eMailerOOo/issues/4>
[93]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/requirements.txt>
[94]: <https://peps.python.org/pep-0508/>
[95]: <https://prrvchr.github.io/eMailerOOo/README_fr#pr%C3%A9requis>
[96]: <https://bugs.documentfoundation.org/show_bug.cgi?id=159988>
[97]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/hsqldb>
[98]: <https://github.com/prrvchr/eMailerOOo/issues/7>
[99]: <https://pypi.org/project/decorator/>
[100]: <https://pypi.org/project/ijson/>
[101]: <https://pypi.org/project/packaging/>
[102]: <https://pypi.org/project/setuptools/>
[103]: <https://github.com/prrvchr/eMailerOOo/security/dependabot/1>
[104]: <https://pypi.org/project/validators/>
[105]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/Options.xcu>
[106]: <https://bugs.documentfoundation.org/show_bug.cgi?id=164040>
[107]: <https://github.com/LibreOffice/loeclipse/pull/123>
[108]: <https://ant.apache.org/>
[109]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/build.xml>
[110]: <https://pypi.org/project/six/>
[111]: <https://github.com/prrvchr/eMailerOOo/issues/10>
[112]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/AsyncCallback.html>
