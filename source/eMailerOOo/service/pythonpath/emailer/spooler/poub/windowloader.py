import uno
import unohelper
from com.sun.star.awt import XWindowListener, XCallback

from time import sleep

class WindowLoader(unohelper.Base, XWindowListener, XCallback):
    def __init__(self, ctx, callback):
        self._ctx = ctx
        self._callback = callback
        self._loaded = False
        
        # Service asynchrone pour le peuplement lourd
        self._callback_service = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.AsyncCallback", ctx)

    # ──── INVOCATION DE L'ÉVÉNEMENT (XWindowListener) ────
    
    def windowShown(self, event):
        """Déclenché au moment PRÉCIS où le dialogue devient visible à l'écran."""
        if not self._loaded:
            self._loaded = True
            # La fenêtre est visible. On planifie immédiatement le remplissage lourd 
            # de manière asynchrone pour ne pas geler le rafraîchissement initial de l'UI.
            self._callback_service.addCallback(self, None)

    def windowHidden(self, event):
        pass

    def windowMoved(self, event):
        pass

    def windowResized(self, event):
        pass

    def disposing(self, event):
        pass

    # ──── REMPLISSAGE ASYNCHRONE (XCallback / XAsyncCall) ────

    def notify(self, dummy_event):
        """Méthode de traitement de vos données lourdes (ex: requêtes HSQLDB)."""
        try:
            for i in range(15):
                sleep(1)
                print("WindowLoader.notify() i: %s" % i)

            self._callback()
            print("[eMailerOOo] Fenêtre peuplée de manière asynchrone avec succès !")
            
        except Exception as e:
            print(f"Erreur lors du peuplement de la fenêtre : {e}")
