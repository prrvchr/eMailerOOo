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

from com.sun.star.logging.LogLevel import INFO
from com.sun.star.logging.LogLevel import SEVERE

from .setupdatabase import SetupDataBase

from ..cancel import Cancel

from ..runner import Extension
from ..runner import Java
from ..runner import Pip
from ..runner import Python

from ...unotool import deregisterStartupJob
from ...unotool import getConfiguration
from ...unotool import getPathSubstitution
from ...unotool import getResourceLocation
from ...unotool import getStringResource
from ...unotool import saveWindowPosition

from ...logger import getLogger

from ...configuration import g_basename
from ...configuration import g_defaultlog
from ...configuration import g_extension
from ...configuration import g_identifier

import traceback


class SetupModel():
    def __init__(self, ctx, name, code, database, java, agent, python, extensions):
        self._ctx = ctx
        self._status = 0
        self._name = name
        self._code = code
        self._java = java
        self._checkJava = None
        self._agent = agent
        self._database = database
        self._python = python
        self._checkPython = None
        self._checkPip = None
        self._extensions = extensions
        self._checkExtension = None
        self._dependencies = ()
        self._modules = []
        self._cancel = Cancel()
        self._extension = {'title': g_extension}
        self._position = 'SetupPosition'
        self._program = getPathSubstitution(ctx, '$(prog)')
        self._pythonpath = getResourceLocation(self._ctx, g_identifier, 'service/pythonpath')
        self._logger = getLogger(ctx, g_defaultlog, g_basename)
        self._config = getConfiguration(ctx, g_identifier, True)
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'SetupWindow')
        self._resources = {'Title': 'SetupWindow.Title',
                           'Header': 'SetupWindow.Label1.Label.%s',
                           'Dependency': 'SetupWindow.Label2.Label',
                           'Java': 'SetupWindow.Label5.Label',
                           'Text': 'SetupWindow.Label15.Label',
                           'Install': 'SetupWindow.Label18.Label'}

    def close(self):
        self._cancel.set()

    def savePosition(self, position):
        saveWindowPosition(self._config, position, self._position)

    def getCancel(self):
        return self._cancel

    def getPage(self, step, **kwargs):
        return step, self._getHeader(step, **kwargs)

    def getViewData(self):
        return self._getDialogPosition(), self._getTitle(), self._getHeader(1, **self._extension)

    def getCheckExtension(self, callback):
        self._checkExtension = Extension(self._ctx, callback, self._extensions)
        return self._checkExtension

    def getCheckJava(self, callback):
        self._checkJava = Java(self._ctx, callback, self._java, self._agent)
        return self._checkJava

    def getCheckDataBase(self, callback, progress):
        self._checkDataBase = SetupDataBase(self._ctx, callback, progress, self._database)
        return self._checkDataBase

    def getCheckPython(self, callback):
        self._checkPython = Python(self._ctx, callback, self._python)
        return self._checkPython

    def getCheckPip(self, callback):
        self._checkPip = Pip(self._ctx, callback, self._python, self._checkPython.getModules())
        return self._checkPip

    def hasDataBase(self):
        return self._database is not None

    def hasJava(self):
        return self._java is not None

    def hasPython(self):
        return self._python is not None

    def hasExtensions(self):
        return len(self._extensions) != 0 

    def deregisterJob(self):
        deregisterStartupJob(self._ctx, self._name)

    def getExtensionResult(self):
        return self._checkExtension.getResult()

    def getJavaResult(self):
        return self._checkJava.getResult()

    def getPythonResult(self):
        return self._checkPython.getResult()

    def getPipResult(self):
        return self._checkPip.getResult()

    def _getJavaResult(self, code, minimum, version):
        return self._getJavaMessage(self._checkJava.getJavaVersion())

    def _getJavaMessage(self, version):
        return 'Java JDK version: ' + version

    def _log(self, level, code, *args):
        self._logger.logprb(level, 'SetupModel', 'installModules()', code, *args)

    def _getDialogPosition(self):
        return uno.createUnoStruct('com.sun.star.awt.Point', *self._config.getByName(self._position))

# SetupModel StringRessoure methods
    def _getTitle(self):
        return self._resolver.resolveString(self._resources.get('Title')).format(**self._extension)

    def _getHeader(self, step, **kwargs):
        resource = self._resources.get('Header') % step
        return self._resolver.resolveString(resource).format(**kwargs)

    def _getDependencyText(self, data):
        return self._resolver.resolveString(self._resources.get('Dependency')) % data

    def _getJavaText(self, version=''):
        return self._resolver.resolveString(self._resources.get('Java')) + version

    def _getProgressText(self, *args):
        return self._resolver.resolveString(self._resources.get('Text')) % args

    def _getInstallText(self, module):
        return self._resolver.resolveString(self._resources.get('Install')) + module

