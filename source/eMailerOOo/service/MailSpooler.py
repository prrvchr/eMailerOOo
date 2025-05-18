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

import unohelper

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from com.sun.star.uno import Exception as UNOException

from emailer import Spooler

from emailer import getLogger

from emailer import g_identifier
from emailer import g_spoolerlog

from threading import Lock
import traceback

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationName = 'io.github.prrvchr.eMailerOOo.MailSpooler'
g_ServiceNames = ('com.sun.star.mail.MailSpooler', )


class MailSpooler():
    def __new__(cls, ctx, *args, **kwargs):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    logger = getLogger(ctx, g_spoolerlog)
                    try:
                        cls._instance = Spooler(ctx, cls._lock, logger, g_ImplementationName)
                        logger.logprb(INFO, 'MailSpooler', '__new__', 1001, g_ImplementationName)
                    except UNOException as e:
                        if cls._logger is None:
                            cls._logger = logger
                        logger.logprb(SEVERE, 'MailSpooler', '__new__', 1002, g_ImplementationName, e.Message)
                        return None
        return cls._instance

    # XXX: If the spooler fails to load then we keep a reference
    # XXX: to the logger so we can read the error message later
    _logger = None
    _instance = None
    _lock = Lock()

g_ImplementationHelper.addImplementation(MailSpooler,                     # UNO object class
                                         g_ImplementationName,            # Implementation name
                                         g_ServiceNames)                  # List of implemented services
