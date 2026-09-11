---
layout: default
title: eMailerOOo historique (Français)
permalink: /change/fr/
redirect_from:
  - /CHANGELOG_fr
  - /CHANGELOG_fr.html
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
# [![eMailerOOo logo][1]][2] Historique

**This [document][3] in English.**

Concernant l'installation, la configuration et l'utilisation, veuillez consulter la **[documentation][4]**.

### Ce qui a été fait pour la version 0.0.1:

- Ecriture de [IspDB][5] ou l'assistant de configuration de connexion aux serveurs SMTP permettant:
    - De trouver les paramètres de connexion à un serveur SMTP à partir d'une adresse courriel. D'ailleur je remercie particulierement Mozilla, pour [Thunderbird autoconfiguration database][6] ou IspDB, qui à rendu ce défi possible...
    - D'afficher l'activité du service UNO `com.sun.star.mail.MailServiceProvider` lors de la connexion au serveur SMTP et l'envoi d'un courriel. 

- Ecriture du [Spouleur][7] de courriels permettant:
    - D'afficher les travaux d'envoi de courriel avec leurs états respectifs.
    - D'afficher l'activité du service UNO `com.sun.star.mail.SpoolerService` lors de l'envoi de courriels.
    - De démarrer et arrêter le service spouleur.

- Ecriture du [Merger][8] ou l'assistant de publipostage de courriels permettant:
    - De créer des listes de diffusions.
    - De fusionner et convertir au format HTML le document courant pour en faire le message du courriel.
    - De fusionner et/ou convertir au format PDF d'éventuel fichiers joints au courriel.

- Ecriture du [Mailer][9] de document permettant:
    - De convertir au format HTML le document pour en faire le message du courriel.
    - De convertir au format PDF d'éventuel fichiers joints au courriel.

- Ecriture d'un [Grid][10] piloté par un `com.sun.star.sdb.RowSet` permettant:
    - D'être paramètrable sur les colonnes à afficher.
    - D'être paramètrable sur l'ordre de tri à afficher.
    - De sauvegarder les paramètres d'affichage.

### Ce qui a été fait pour la version 0.0.2:

- Réécriture de [IspDB][5] ou Assistant de configuration de connexion aux serveurs de messagerie afin d'intégrer la configuration de la connexion IMAP.
    - Utilisation de [IMAPClient][11] version 2.2.0: une bibliothèque cliente IMAP complète, Pythonic et facile à utiliser.
    - Extension des fichiers IDL [com.sun.star.mail.*][12]:
        - [XMailMessage2.idl][13] prend désormais en charge la hiérarchisation des courriels (thread).
        - La nouvelle interface [XImapService][14] permet d'accéder à une partie de la bibliothèque IMAPClient.

- Réécriture du [Spouleur][7] afin d'intégrer des fonctionnalités IMAP comme la création d'un fil récapitulant le publipostage et regroupant tous les courriels envoyés.

- Soumission de l'extension eMailerOOo à Google et obtention de l'autorisation d'utiliser son API GMail afin d'envoyer des courriels avec un compte Google.

### Ce qui a été fait pour la version 0.0.3:

- Réécriture du [Grid][10] afin de permettre:
    - Le tri sur une colonne avec l'intégration du service UNO [SortableGridDataModel][15].
    - La génération des filtres des enregistrements nécessaires au service [Spouleur][7].
    - Le partage avec le module python [Grid][16] de l'extension [jdbcDriverOOo][17].

- Réécriture du [Merger][8] afin de permettre:
    - La gestion du nom du Schema dans de nom des tables afin d'être compatible avec la version 0.0.4 de [jdbcDriverOOo][17].
    - La création de liste de diffusion sur un groupe du carnet d'adresse et permettant de suivre la modification de son contenu.
    - L'utilisation de clé primaire, qui peuvent être composite, supportant les [DataType][18] `VARCHAR` et `INTEGER` ou derivé.
    - Un aperçu du document avec des champs de fusion remplis plus rapidement grâce au [Grid][10].

- Réécriture du [Spouleur][7] afin de permettre:
    - L'utilisation des nouveaux filtres supportant les clés primaires composite fourni par le [Merger][8].
    - L'utilisation du nouveau [Grid][10] permettant le tri sur une colonne.

- Encore plein d'autres choses...

### Ce qui a été fait pour la version 1.0.0:

- L'extension **smtpMailerOOo** a été renomé en **eMailerOOo**.

### Ce qui a été fait pour la version 1.0.1:

- L'absence ou l'obsolescence des extensions **OAuth2OOo** et/ou **jdbcDriverOOo** nécessaires au bon fonctionnement de **eMailerOOo** affiche désormais un message d'erreur. Ceci afin d'éviter qu'un dysfonctionnement tel que le [dysfonctionnement #3][19] ne se reproduise...

- La base de données HsqlDB sous-jacente peut être ouverte dans Base avec: **Outils -> Options -> Internet -> eMailerOOo -> Base de données**.

- Le menu **Outils -> Add-ons** s'affiche désormais correctement en fonction du contexte.

- Encore plein d'autres choses...

### Ce qui a été fait pour la version 1.0.2:

- Si aucune configuration n'est trouvée dans l'assistant de configuration de la connexion (IspDB Wizard) alors il est possible de configurer la connexion manuellement. Voir [dysfonctionnement #5][20].

### Ce qui a été fait pour la version 1.1.0:

- Dans l'assistant de configuration de la connexion (IspDB Wizard) il est maintenant possible de désactiver la configuration IMAP.  
    En conséquence, cela n'envoie plus de fil de discussion (message IMAP) lors de la fusion d'un mailing.  
    Dans ce même assistant, il est désormais possible de saisir une adresse courriel de réponse.

- Dans l'assistant de fusion d'email, il est désormais possible d'insérer des champs de fusion dans l'objet du courriel. Voir [dysfonctionnement #6][21].  
    Dans le sujet d'un courriel, un champ de fusion est composé d'une accolade ouvrante, du nom de la colonne référencée (sensible à la casse) et d'une accolade fermante (ie: `{NomDeLaColonne}`).  
    Lors de la saisie du sujet du courriel, une erreur de syntaxe dans un champ de fusion sera signalée et empêchera la soumission du mailing.

- Il est désormais possible dans le Spouleur de visualiser les courriels au format eml.

- Un service [com.sun.star.mail.MailUser][22] permet désormais d'accéder à une configuration de connexion (SMTP et/ou IMAP) depuis une adresse courriel qui suite la rfc822.  
    Un autre service [com.sun.star.datatransfer.TransferableFactory][23] permet, comme son nom l'indique, la création de [Transferable][24] à partir d'un texte (string), d'une séquence binaire, d'une Url (file://...) ou un flux de données (InputStream).  
    Ces deux nouveaux services simplifient grandement l'API mail de LibreOffice et permettent d'envoyer des courriels depuis Basic. Voir le [dysfonctionnement #4][25].  
    Vous trouverez une macro Basic vous permettant d'envoyer des emails dans : **Outils -> Macros -> Editer les Macros... -> eMailerOOo -> SendEmail**.

### Ce qui a été fait pour la version 1.1.1:

- Prise en charge de la version 1.2.0 de l'extension **OAuth2OOo**. Les versions précédentes ne fonctionneront pas avec l'extension **OAuth2OOo** 1.2.0 ou ultérieure.

### Ce qui a été fait pour la version 1.2.0:

- Tous les paquets Python nécessaires à l'extension sont désormais enregistrés dans un fichier [requirements.txt][26] suivant la [PEP 508][27].
- Désormais si vous n'êtes pas sous Windows alors les paquets Python nécessaires à l'extension peuvent être facilement installés avec la commande:  
  `pip install requirements.txt`
- Modification de la section [Prérequis][28].

### Ce qui a été fait pour la version 1.2.1:

- Correction d'une régression permettant l'affichage des erreurs dans le Spouleur.
- Intégration d'un correctif pour contourner le [dysfonctionnement #159988][29].

### Ce qui a été fait pour la version 1.2.2:

- La création de la base de données, lors de la première connexion, utilise l'API UNO proposée par l'extension jdbcDriverOOo depuis la version 1.3.2. Cela permet d'enregistrer toutes les informations nécessaires à la création de la base de données dans 5 tables texte qui sont en fait [5 fichiers csv][30].
- L'extension vous demandera d'installer les extensions OAuth2OOo et jdbcDriverOOo en version respectivement 1.3.4 et 1.3.2 minimum.
- De nombreuses corrections.

### Ce qui a été fait pour la version 1.2.3:

- Correction d'une régression provenant de la version 1.2.2 et empêchant la soumission des travaux dans le spouleur de courriels.
- Correction du [dysfonctionnement #7][31] ne permettant pas l'affichage des messages d'erreur en cas de configuration incorrecte.

### Ce qui a été fait pour la version 1.2.4:

- Mise à jour du paquet [Python decorator][32] vers la version 5.1.1.
- Mise à jour du paquet [Python ijson][33] vers la version 3.3.0.
- Mise à jour du paquet [Python packaging][34] vers la version 24.1.
- Mise à jour du paquet [Python setuptools][35] vers la version 72.1.0 afin de répondre à l'[alerte de sécurité Dependabot][36].
- Mise à jour du paquet [Python validators][37] vers la version 0.33.0.
- L'extension vous demandera d'installer les extensions OAuth2OOo et jdbcDriverOOo en version respectivement 1.3.6 et 1.4.2 minimum.

### Ce qui a été fait pour la version 1.2.5:

- Mise à jour du paquet [Python setuptools][35] vers la version 73.0.1.
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

    Cela devrait permettre d'implémenter une API tierce pour l'envoi de courriels simplement en modifiant le fichier de configuration [Options.xcu][38].
- Pour fonctionner, ces nouvelles fonctionnalités nécessitent l'extension OAuth2OOo en version 1.3.9 minimum.
- La commande permettant d'ouvrir un courriel dans Thunderbird ne peut actuellement être modifiée que dans la configuration de LibreOffice (ie: Outils -> Options... -> Avancé -> Ouvrir la configuration avancée)
- Le non rafraîchissement des barres de défilement dans les listes multicolonnes (ie: grid) a été corrigé et sera disponible à partir de LibreOffice 24.8.4, voir [SortableGridDataModel cannot be notified for changes][39].
- L'ouverture des courriels dans votre navigateur ne fonctionne pas avec un compte Microsoft, l'url permettant cela n'a pas encore été trouvée et il semble que ce ne serait pas possible (ie: le popup doit être ouvert par la fenêtre Outlook)?
- De nombreuses corrections.

### Ce qui a été fait pour la version 1.3.0:

- L'extension vous demandera d'installer les extensions OAuth2OOo et jdbcDriverOOo en version respectivement 1.4.0 et 1.4.6 minimum.
- Seuls les fournisseurs disposant d'une API tierce ou d'une authentification OAuth2 et disposant d'une entrée dans la configuration de LibreOffice proposeront l'authentification OAuth2 par défaut dans l'assistant de configuration de connexion (ie: IspDB Wizard).
- Les fournisseurs de messagerie `yahoo.com` et `aol.com` ont été intégrés. Afin de faciliter la configuration, un lien vers la page permettant la création d'un mot de passe d'application a été ajouté à l'assistant de configuration de connexion. Si vous pensez que des liens vers d'autres fournisseurs manquent, veuillez ouvrir un dysfonctionnement afin que je puisse les rajouter.
- Mise à jour du paquet [Python IMAPClient][11] vers la version 3.0.1.
- Grâce aux améliorations apportées au [plugin Eclipse][40], il est désormais possible de créer le fichier de l'extension en utilisant la ligne de commande et l'outil de création d'archive [Apache Ant][41], voir le fichier [build.xml][42].
- L'extension refusera de s'installer sous OpenOffice quelle que soit la version ou LibreOffice autre que 7.x ou supérieur.
- Ajout des fichiers binaires nécessaires aux bibliothèques Python pour fonctionner sous Linux et LibreOffice 24.8 (ie: Python 3.9).
- De nombreuses corrections.

### Ce qui a été fait pour la version 1.3.1:

- Mise à jour du paquet [Python packaging][34] vers la version 24.2.
- Mise à jour du paquet [Python setuptools][35] vers la version 75.8.0.
- Mise à jour du paquet [Python six][43] vers la version 1.17.0.
- Mise à jour du paquet [Python validators][37] vers la version 0.34.0.
- Support de Python version 3.13.

### Ce qui a été fait pour la version 1.4.0:

- Mise à jour du paquet [Python packaging][34] vers la version 25.0.
- Rétrogradage du paquet [Python setuptools][35] vers la version 75.3.2, afin d'assurer la prise en charge de Python 3.8.
- Déploiement de l'enregistrement passif permettant une installation beaucoup plus rapide des extensions et de différencier les services UNO enregistrés de ceux fournis par une implémentation Java ou Python. Cet enregistrement passif est assuré par l'extension [LOEclipse][44] via les [PR#152][45] et [PR#157][46].
- Modification de [LOEclipse][44] pour prendre en charge le nouveau format de fichier `rdb` produit par l'utilitaire de compilation `unoidl-write`. Les fichiers `idl` ont été mis à jour pour prendre en charge les deux outils de compilation disponibles: idlc et unoidl-write.
- Il est désormais possible de créer le fichier oxt de l'extension eMailerOOo uniquement avec Apache Ant et une copie du dépôt GitHub. La section [Comment créer l'extension][47] a été ajoutée à la documentation.
- Implémentation de [PEP 570][48] dans la [journalisation][49] pour prendre en charge les arguments multiples uniques.
- Pour garantir la création correcte de la base de données eMailerOOo, il sera vérifié que l'extension jdbcDriverOOo a `com.sun.star.sdb` comme niveau d'API.
- Écriture de macros pour pouvoir placer des menus personnalisés où vous le souhaitez. Pour faciliter la création de ces menus personnalisés, la section [Comment personnaliser les menus de LibreOffice][50] a été ajoutée à la documentation.
- Nécessite l'extension **jdbcDriverOOo en version 1.5.0 minimum**.
- Nécessite l'extension **OAuth2OOo en version 1.5.0 minimum**.

### Ce qui a été fait pour la version 1.4.1:

- Dans l'assistant de connexion, si l'adresse courriel donnée n'est pas trouvée dans Mozilla IspDB ou si vous êtes hors ligne, les noms de serveur peuvent être de simples noms d'hôtes et les ports valides s'étendront jusqu'à 65535. Ceci afin de répondre à l'[issue#10][51].
- Les problèmes d'actualisation de la deuxième page de l'assistant de connexion ont été résolus par l'utilisation du service UNO `com.sun.star.awt.AsyncCallback`.
- Nécessite l'extension **jdbcDriverOOo en version 1.5.4 minimum**.
- Nécessite l'extension **OAuth2OOo en version 1.5.1 minimum**.

### Ce qui a été fait pour la version 1.4.2:

- Support de LibreOffice 25.2.x et 25.8.x sous Windows 64 bits.
- Nécessite l'extension **OAuth2OOo en version 1.5.2 minimum**.

### Ce qui a été fait pour la version 1.5.0:

- Modification de l'assistant utilisé lors de la fusion des courriels afin qu'il s'ouvre dans une fenêtre dédiée plutôt que modale comme auparavant.
- Modification également du Spouleur de courriels afin qu'il s'ouvre dans une fenêtre dédiée plutôt que modale comme auparavant.
- Ces deux nouvelles fenêtres affichent maintenant une barre de progression ainsi qu'un indicateur d'état lorsque des tâches en arrière-plan sont lancées.
- Si des tâches sont démarrées alors que ces fenêtres sont demandées à être fermées, alors les tâches seront annulées si possible et leur achèvement sera attendu avant la fermeture.
- Le Spouleur de courriels a été entièrement réécrit. Il propose désormais trois tâches pour envoyer des courriels, afficher un courriel et fusioner un document:
  - [sender.py][52]
  - [mailer.py][53]
  - [viewer.py][54]
- Ajout de l'interface [XTaskEvent.idl][55] à l'API UNO. Cette nouvelle interface, qui est la transcription de la classe Python [threading.Event][56], permet de contrôler une tâche exécutée par le [Dispatcher][57] de LibreOffice.
- Si des fichiers sont joints au courriel et au format PDF, alors ils suivront les paramètres de configuration de LibreOffice qui se trouvent dans: **Fichier -> Exporter vers -> Exporter au format PDF** lors de leur transformation.
- Toutes les méthodes nécessaires à l'affichage et s'exécutant en arrière-plan utilisent désormais le service UNO [com.sun.star.awt.AsyncCallback][58] pour le rappel.
- Si l'extension jdbcDriverOOo fonctionne sans l'instrumentation Java, un message d'avertissement s'affichera dans les options de l'extension.
- De nombreuses corrections et quelques nouveautés que je vous laisse découvrir.
- Nécessite l'extension **jdbcDriverOOo en version 1.6.0 minimum**.
- Nécessite l'extension **OAuth2OOo en version 1.6.0 minimum**.
- A été testé avec LibreOfficeDev 26.2.

### Ce qui a été fait pour la version 1.5.1:

- Toute erreur survenant lors de l'envoi d'un courriel n'affectera pas son statut si elle se produit lors de la préparation et non lors de l'envoi. Cela permet de corriger l'erreur et de réessayer l'envoi.
- Il est possible d'ouvrir une pièce jointe directement depuis la liste des fichiers joints à un courriel.
- Si cette opération est effectuée sur un fichier joint (Writer ou Calc) à fusionner, les champs de fusion de ce document ouvert dans LibreOffice suivront la sélection des Grids `Destinataires disponibles` et/ou `Destinataires sélectionnés`.
- Toutes les fenêtres modales s'ouvrent désormais correctement en mode modal.
- Nécessite l'extension **jdbcDriverOOo en version 1.6.1 minimum**.
- Nécessite l'extension **OAuth2OOo en version 1.6.1 minimum**.

### Ce qui a été fait pour la version 1.5.2:

- Correction d'une régression qui empêchait la soumission de travaux valides au Spouleur en vue d'une fusion.

### Ce qui a été fait pour la version 1.7.0:


### Que reste-t-il à faire pour la version 1.7.0:

- Ajouter de nouvelles langues pour l’internationalisation...

- Tout ce qui est bienvenu...

[1]: </img/emailer.svg#collapse>
[2]: <https://prrvchr.github.io/eMailerOOo/>
[3]: <https://prrvchr.github.io/eMailerOOo/change/>
[4]: <https://prrvchr.github.io/eMailerOOo/fr/>
[5]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/ispdb>
[6]: <https://wiki.mozilla.org/Thunderbird:Autoconfiguration>
[7]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler>
[8]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/merger>
[9]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/mailer>
[10]: <https://github.com/prrvchr/eMailerOOo/tree/master/uno/lib/uno/grid>
[11]: <https://github.com/mjs/imapclient#readme>
[12]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/idl/com/sun/star/mail>
[13]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailMessage2.idl>
[14]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XImapService.idl>
[15]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/grid/SortableGridDataModel.html>
[16]: <https://github.com/prrvchr/jdbcDriverOOo/tree/master/source/jdbcDriverOOo/service/pythonpath/jdbcdriver/grid>
[17]: <https://prrvchr.github.io/jdbcDriverOOo/README_fr>
[18]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/sdbc/DataType.html>
[19]: <https://github.com/prrvchr/eMailerOOo/issues/3>
[20]: <https://github.com/prrvchr/eMailerOOo/issues/5>
[21]: <https://github.com/prrvchr/eMailerOOo/issues/6>
[22]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/mail/XMailUser.idl>
[23]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/idl/com/sun/star/datatransfer/XTransferableFactory.idl>
[24]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/datatransfer/XTransferable.html>
[25]: <https://github.com/prrvchr/eMailerOOo/issues/4>
[26]: <https://github.com/prrvchr/eMailerOOo/releases/latest/download/requirements.txt>
[27]: <https://peps.python.org/pep-0508/>
[28]: <https://prrvchr.github.io/eMailerOOo/README_fr#pr%C3%A9requis>
[29]: <https://bugs.documentfoundation.org/show_bug.cgi?id=159988>
[30]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/hsqldb>
[31]: <https://github.com/prrvchr/eMailerOOo/issues/7>
[32]: <https://pypi.org/project/decorator/>
[33]: <https://pypi.org/project/ijson/>
[34]: <https://pypi.org/project/packaging/>
[35]: <https://pypi.org/project/setuptools/>
[36]: <https://github.com/prrvchr/eMailerOOo/security/dependabot/1>
[37]: <https://pypi.org/project/validators/>
[38]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/Options.xcu>
[39]: <https://bugs.documentfoundation.org/show_bug.cgi?id=164040>
[40]: <https://github.com/LibreOffice/loeclipse/pull/123>
[41]: <https://ant.apache.org/>
[42]: <https://github.com/prrvchr/eMailerOOo/blob/master/source/eMailerOOo/build.xml>
[43]: <https://pypi.org/project/six/>
[44]: <https://github.com/LibreOffice/loeclipse>
[45]: <https://github.com/LibreOffice/loeclipse/pull/152>
[46]: <https://github.com/LibreOffice/loeclipse/pull/157>
[47]: <https://prrvchr.github.io/eMailerOOo/README_fr#comment-cr%C3%A9er-lextension>
[48]: <https://peps.python.org/pep-0570/>
[49]: <https://github.com/prrvchr/eMailerOOo/blob/master/uno/lib/uno/logger/logwrapper.py#L106>
[50]: <https://prrvchr.github.io/eMailerOOo/README_fr#comment-personnaliser-les-menus-de-libreoffice>
[51]: <https://github.com/prrvchr/eMailerOOo/issues/10>
[52]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/thread/sender.py>
[53]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/thread/mailer.py>
[54]: <https://github.com/prrvchr/eMailerOOo/tree/master/source/eMailerOOo/service/pythonpath/emailer/spooler/thread/viewer.py>
[55]: <https://github.com/prrvchr/eMailerOOo/blob/master/uno/rdb/idl/com/sun/star/task/XTaskEvent.idl>
[56]: <https://docs.python.org/3/library/threading.html#threading.Event>
[57]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XDispatch.html#dispatch>
[58]: <https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/AsyncCallback.html>
