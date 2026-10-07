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

from com.sun.star.awt import XCallback

from ..unotool import getCallBack
from ..unotool import getResourceLocation
from ..unotool import getSimpleFile

from .helper import checkAgent
from .helper import checkExtension
from .helper import checkPython
from .helper import getJavaStatus
from .helper import getJavaVersion

from ..configuration import g_identifier

from threading import Timer
import traceback


class CheckSetup(unohelper.Base,
                 XCallback):
    def __init__(self, ctx, callback, database=None, java=None, agent=None, python=False, **extensions):
        self._ctx = ctx
        self._callback = callback
        self._requirements = 'requirements.txt'
        self._asyncCall = getCallBack(ctx)
        self._results = self._getDefaultResults(True)
        self._count = 0
        self._steps = self._getCheckStep(ctx, database, java, agent, python, extensions)

    def start(self):
        Timer(0.1, self._call).start()

    def notify(self, data):
        try:
            self._count += 1
            call, *args = next(self._steps)
            call(*args)
            Timer(0.1, self._call).start()
        except StopIteration:
            self._callback(**self._results)
        except Exception as e:
            print("CheckSetup.notify() count: %s - ERROR: %s" % (self._count, traceback.format_exc()))
            self._callback(**self._getDefaultResults(False))

    def stepCheckExtension(self, identifier, infos):
        self._results['extension'] &= checkExtension(self._ctx, identifier, *infos)

    def stepCheckJavaStatus(self, agent, extension, version, script):
        self._code = getJavaStatus(self._ctx)
        self._results['java'] = self._hasJava()

    def stepCheckJavaVersion(self, agent, extension, version, script):
        self._code, _, _ = getJavaVersion(self._ctx, extension, version, script)
        self._results['java'] = self._hasJava()

    def stepCheckJavaAgent(self, service, url, agent):
        self._results['agent'] = checkAgent(self._ctx, extension, version, script)

    def stepCheckDataBase(self, database):
        self._results['database'] = getSimpleFile(self._ctx).exists(database)

    def stepGetRequirementUrl(self, ctx):
        self._url = getResourceLocation(ctx, g_identifier, self._requirements)

    def stepCheckPython(self):
        self._results['python'] = checkPython(self._url)

    def _getCheckStep(self, ctx, database, java, agent, python, extensions):
        if extensions:
            for identifier, infos in extensions:
                yield self._getStepCheckExtension(identifier, infos)
        if java:
            yield self._getStepCheckJavaStatus(*java)
            if self._hasJava():
                yield self._getStepCheckJavaStatus(*java)
            if self._hasJava() and agent:
                yield self._getStepCheckJavaAgent(*agent)
        if database:
            yield self._getStepCheckDataBase(database)
        if python:
            yield self._getStepGetRequirementUrl(ctx)
            if getSimpleFile(self._ctx).exists(self._url):
                yield self._getStepCheckPython()

    def _call(self):
       self._asyncCall.addCallback(self, None)

    def _getStepCheckExtension(self, identifier, infos):
        return self.stepCheckExtension, identifier, infos

    def _getStepCheckJavaStatus(self, agent, extension, version, script):
        return self.stepCheckJavaStatus, agent, extension, version, script

    def _getStepCheckJavaVersion(self, agent, extension, version, script):
        return self.stepCheckJavaVersion, agent, extension, version, script

    def _getStepCheckJavaAgent(self, service, url, agent):
        return self.stepCheckJavaAgent, service, url, agent

    def _getStepCheckDataBase(self, database):
        return self.stepCheckDataBase, database

    def _getStepGetRequirementUrl(self, ctx):
        return self.stepGetRequirementUrl, ctx

    def _getStepCheckPython(self):
        return self.stepCheckPython,

    def _getDefaultResults(self, value):
        return {'extension': value, 'java': value, 'agent':value, 'database': value, 'python': value}

    def _hasJava(self):
        return self._code == 0

    @classmethod
    def isValid(cls, ctx, database=None, java=None, python=False, **extensions):
        code = None
        for identifier, infos in extensions:
            if not checkExtension(ctx, identifier, *infos):
                return False
        if java:
            code = getJavaStatus(ctx)
            if code > 0:
                return False
            code, _, _ = getJavaVersion(ctx, *java)
            if code > 0:
                return False
        if database and not getSimpleFile(ctx).exists(database):
            return False
        if python:
            url = getResourceLocation(ctx, g_identifier, self._requirements)
            if getSimpleFile(self._ctx).exists(url):
                return checkPython(url)
            return False
        return True

