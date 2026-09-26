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

from .check import Check

from ..helper import checkAgent
from ..helper import getJavaStatus
from ..helper import getJavaVersion

import traceback


class Java(Check):
    def __init__(self, ctx, callback, java, agent=None):
        super().__init__(ctx, callback)
        self._code = 4
        self._minimum = ''
        self._version = ''
        self.total = 1
        self.label1 = self.resolver.resolveString(201)
        self.label2 = self.resolver.resolveString(202)
        self.steps = self._getCheckStep(java, agent)

    def getResult(self):
        return self._minimum

    def getJavaVersion(self):
        return self._version if self._hasJavaVersion() else '...'

    def callback(self, success):
        self._callback(self._getSuccess(success), self._code, self._minimum, self._version)

    def stepCheckJavaStatus(self, agent, extension, version, script):
        self._minimum = version
        self._code = getJavaStatus(self._ctx)

    def stepCheckJavaVersion(self, agent, extension, version, script):
        self._code, _, self._version = getJavaVersion(self._ctx, extension, version, script)

    def stepCheckJavaAgent(self, service, url, agent):
        self._agent = checkAgent(self._ctx, extension, version, script)

    def _getCheckStep(self, java, agent):
        yield self._getStepCheckJavaStatus(agent, *java)
        if not self._hasJava():
            return
        self.total += 1
        yield self._getStepCheckJavaVersion(agent, *java)
        if not self._hasJava() or agent is None:
            return
        self.total += 1
        yield self._getStepCheckJavaAgent(*agent)

    def _getStepCheckJavaStatus(self, agent, extension, version, script):
        return 211, (), self.stepCheckJavaStatus, agent, extension, version, script

    def _getStepCheckJavaVersion(self, agent, extension, version, script):
        return 221, (version, ), self.stepCheckJavaVersion, agent, extension, version, script

    def _getStepCheckJavaAgent(self, service, url, agent):
        return 231, (), self.stepCheckJavaAgent, service, url, agent

    def _getSuccess(self, success):
        return success and self._hasJava()

    def _hasJavaVersion(self):
        return self._code < 2

    def _hasJava(self):
        return self._code == 0

