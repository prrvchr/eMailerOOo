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

from ...unotool import getPathSubstitution
from ...unotool import getResourceLocation
from ...unotool import getSimpleFile

from ..helper import getPipCommand
from ..helper import getStartupInfo
from ..helper import pipInstall

import os
import traceback


class Pip(Check):
    def __init__(self, ctx, callback, identifier, modules):
        super().__init__(ctx, callback)
        self._installed = []
        self._aborted = []
        self._error = None
        self._command = None
        self._info = None
        self.total = len(modules) + 1
        self.label1 = self.resolver.resolveString(401)
        self.label2 = self.resolver.resolveString(402)
        self.steps = self._getCheckStep(identifier, modules)

    def getResult(self):
        if len(self._aborted):
            return ', '.join(self._aborted)
        return ', '.join(self._installed)

    def callback(self, success):
        self._callback(self._getSuccess(success))

    def stepSetPipCommand(self, identifier):
        isWindows = os.name == 'nt'
        program = getPathSubstitution(self._ctx, '$(prog)')
        python = getResourceLocation(self._ctx, identifier, 'service/pythonpath')
        self._command = getPipCommand(program, isWindows, python)
        self._info = getStartupInfo(isWindows)

    def stepInstallPackage(self, package):
        error = pipInstall(self._info, self._command, package)
        if error is not None:
            self._aborted.append(package)
            self._error = error
        else:
            self._installed.append(package)

    def _getCheckStep(self, identifier, packages):
        yield self._getStepSetPipCommand(identifier)
        for package in packages:
            yield self._getStepInstallPackage(package)

    def _getStepSetPipCommand(self, identifier):
        return 411, (), self.stepSetPipCommand, identifier

    def _getStepInstallPackage(self, package):
        return 421, (package, ), self.stepInstallPackage, package

    def _getSuccess(self, success):
        return success and len(self._aborted) == 0

