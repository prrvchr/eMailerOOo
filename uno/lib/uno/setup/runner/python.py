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

from .check import Check

from ...unotool import getPathSubstitution
from ...unotool import getResourceLocation
from ...unotool import getSimpleFile

from ..helper import checkPython
from ..helper import getPipCommand
from ..helper import getPackageCount
from ..helper import getStartupInfo
from ..helper import pipInstall

import traceback


class Python(Check):
    def __init__(self, ctx, callback, identifier):
        super().__init__(ctx, callback)
        self._installed = []
        self._missing = []
        self._url = None
        self._requirements = 'requirements.txt'
        self.total = 1
        self.label1 = self.resolver.resolveString(301)
        self.label2 = self.resolver.resolveString(302)
        self.steps = self._getCheckStep(identifier)

    def getModules(self):
        if len(self._missing):
            return self._missing
        return self._installed

    def getResult(self):
        if len(self._missing):
            return ', '.join(self._missing)
        return ', '.join(self._installed)

    def callback(self, success):
        self._callback(self._getSuccess(success))

    def stepGetRequirementUrl(self, identifier):
        self._url = getResourceLocation(self._ctx, identifier, self._requirements)

    def stepGetPackageCount(self):
        self.total += getPackageCount(self._url)

    def stepAddInstalled(self, module):
        self._installed.append(module)

    def stepAddMissing(self, module):
        self._missing.append(module)

    def _getCheckStep(self, identifier):
        yield self._getStepGetRequirementUrl(identifier)
        if not getSimpleFile(self._ctx).exists(self._url):
            return
        self.total += 1
        yield self._getStepGetPackageCount()
        yield from checkPython(self._url, self._getStepAddInstalled, self._getStepAddMissing)

    def _getStepGetRequirementUrl(self, identifier):
        return 311, (), self.stepGetRequirementUrl, identifier

    def _getStepGetPackageCount(self):
        return 321, (), self.stepGetPackageCount

    def _getStepAddInstalled(self, module):
        return 331, (module, ), self.stepAddInstalled, module

    def _getStepAddMissing(self, module):
        return 341, (module, ), self.stepAddMissing, module

    def _getSuccess(self, success):
        return success and len(self._missing) == 0

