import uno
import unohelper

from com.sun.star.awt import XCallback

from threading import Timer
from time import sleep


class FluidProgressMock(unohelper.Base,
                        XCallback):
    def __init__(self, ctx, status_indicator, callback):
        self._ctx = ctx
        self._callback = callback
        self._status_indicator = status_indicator
        
        # On veut simuler 15 étapes (15 secondes de travail au total)
        self.TOTAL_STEPS = 15
        self.current_step = 0
        
        # On récupère le service asynchrone pour exécuter le code sur le thread de l'UI
        self._callback_service = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.AsyncCallback", ctx
        )

    def start(self):
        """Déclenche la maquette."""
        if self._status_indicator:
            self._status_indicator.start("Chargement de la fenêtre...", self.TOTAL_STEPS)
        
        # On lance immédiatement la première étape
        self._callback_service.addCallback(self, None)

    def notify(self, dummy_event):
        """
        Cette méthode s'exécute instantanément sur le thread principal.
        Elle NE DOIT PAS contenir de boucle d'attente.
        """
        # 1. On fait avancer la jauge d'une unité
        self.current_step += 1
        
        if self._status_indicator:
            self._status_indicator.setText(f"Étape {self.current_step} / {self.TOTAL_STEPS} complétée...")
            self._status_indicator.setValue(self.current_step)

        # 2. Vérification de fin
        if self.current_step >= self.TOTAL_STEPS:
            if self._status_indicator:
                self._status_indicator.end()
            print("[eMailerOOo Mock] Fin du chargement fluide !")
            self._callback()
            return

        # 3. 🚀 LE SECRET DE LA PAUSE FLUIDE : 
        # On demande à Python d'attendre 1 seconde en tâche de fond (sans bloquer l'UI),
        # puis d'appeler addCallback pour déclencher l'étape suivante.
        def schedule_next():
            #sleep(1)
            self._callback_service.addCallback(self, None)
            
        # Timer non-bloquant : il attend 1.0 seconde et exécute 'schedule_next'
        Timer(1.0, schedule_next).start()
