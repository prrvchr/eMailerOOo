#!
# -*- coding: utf-8 -*-

"""
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
"""

import uno
import unohelper

from com.sun.star.awt import XCallback

import time
import traceback


class SpoolerTask(unohelper.Base,
                  XCallback):
    def __init__(self, ctx, indicator, callback):
        self._ctx = ctx
        self._indicator = indicator
        self._callback = callback
        
        # Simulation de 5 lots de fusion (batches) issus de votre vue HSQLDB
        self.batches = [
            {"id": 1, "name": "Lot Clients Région Nord", "mails_count": 12},
            {"id": 2, "name": "Lot Clients Région Sud", "mails_count": 8},
            {"id": 3, "name": "Lot Partenaires", "mails_count": 15},
            {"id": 4, "name": "Lot Fournisseurs", "mails_count": 5},
            {"id": 5, "name": "Lot Relances Factures", "mails_count": 20}
        ]
        
        # Calcul du total d'e-mails pour la barre de progression
        self.total_mails = sum(b["mails_count"] for b in self.batches)
        self.current_progress = 0
        self.iterator = iter(self.batches)
        self.is_cancelled = False
        
        # Récupération du service de callback asynchrone de LibreOffice
        self.callback_service = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.AsyncCallback", ctx
        )

    def start(self):
        """Déclenche la maquette et initialise le StatusIndicator."""
        if self._indicator:
            # Initialise la barre de progression avec le texte et le maximum
            self._indicator.start("eMailerOOo : Fusion et envoi en cours...", self.total_mails)
        
        # On demande à LibreOffice de planifier le premier lot dès qu'il est disponible
        self.callback_service.addCallback(self, None)

    def notify(self, dummy_event):
        """Méthode appelée par LibreOffice à chaque fois que la main lui est rendue."""
        # 1. Vérification si l'utilisateur a cliqué sur "Annuler" dans l'UI
        if self.is_cancelled:
            self._finalize(reason="Annulé par l'utilisateur")
            return

        # 2. Récupération du lot suivant
        try:
            current_batch = next(self.iterator)
        except StopIteration:
            # Plus aucun lot, le spooler a terminé avec succès
            self._finalize(reason="Succès")
            return

        # 3. SIMULATION DU TRAVAIL LOURD (Fusion de fichiers / Envoi SMTP)
        # On simule un traitement proportionnel au nombre d'e-mails dans le lot
        for _ in range(current_batch["mails_count"]):
            time.sleep(1) # Simule le temps de fusion d'un document individuel
        
        # 4. MISE À JOUR DE L'INTERFACE GRAPHIQUE
        self.current_progress += current_batch["mails_count"]
        if self._indicator:
            # On met à jour le texte pour montrer le lot actuel
            self._indicator.setText(f"Traitement : {current_batch['name']}")
            self._indicator.setValue(self.current_progress)

        # 5. RAPPEL ASYNCHRONE
        # En rappelant addCallback, on dit à LibreOffice : "J'ai fini ce lot. prends 
        # le contrôle quelques millisecondes pour rafraîchir l'UI, puis rappelle-moi 
        # pour le lot suivant."
        self.callback_service.addCallback(self, None)

    def cancel(self):
        """Permet à votre fenêtre de notifier le spooler d'un arrêt d'urgence."""
        self.is_cancelled = True

    def _finalize(self, reason):
        try:
            """Nettoyage final et fermeture de l'indicateur."""
            if self._indicator:
                self._indicator.setText(f"Statut final : {reason}")
                self._indicator.end()
            print(f"[eMailerOOo Mock] Fin du spooler. Raison : {reason}")
            self._callback()
        except Exception as e:
            print("SpoolerTask._finalize() ERROR: %s" % traceback.format_exc())


