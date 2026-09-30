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

from ..helper import getPackageData
from ..helper import getPackageUrl
from ..helper import installPackage

from ...unotool import getResourceLocation

from ...configuration import g_identifier

import traceback


class Pypi(Check):
    def __init__(self, ctx, callback, modules):
        super().__init__(ctx, callback)
        self._installed = {}
        self._aborted = {}
        self._error = None
        self._data = None
        self._url = None
        self.total = 1 + len(modules) * 3
        self.label1 = self.resolver.resolveString(411)
        self.label2 = self.resolver.resolveString(412)
        self.steps = self._getCheckStep(modules)

    def getHeader(self, **kwargs):
        return self.resolver.resolveString(421).format(**kwargs)

    def getResults(self, success):
        return self._getHeader(), self._getResult(), True

    def callback(self, success):
        self._callback(self._getSuccess(success))

    def stepGetPythonPathUrl(self):
        self._path = getResourceLocation(self._ctx, g_identifier, 'service/pythonpath')

    def stepGetPackageData(self, package):
        self._data = getPackageData(package)

    def stepGetPackageUrl(self, package, data):
        self._url, self._name, self._version = getPackageUrl(package, data)

    def stepInstallPackage(self, package, url, name, version, path):
        if installPackage(url, path):
            self._installed[package] = version
        else:
            self._aborted[package] = version

    def _getCheckStep(self, packages):
        yield self._getStepGetPythonPathUrl()
        if self._path:
            for package in packages:
                yield self._getStepGetPackageData(package)
                if self._data:
                    yield self._getStepGetPackageUrl(package, self._data)
                    if self._url:
                        yield self._getStepInstallPackage(package, self._url, self._name, self._version, self._path)

    def _getStepGetPythonPathUrl(self):
        return 431, (), self.stepGetPythonPathUrl

    def _getStepGetPackageData(self, package):
        return 441, (package, ), self.stepGetPackageData, package

    def _getStepGetPackageUrl(self, package, data):
        return 451, (package, ), self.stepGetPackageUrl, package, data

    def _getStepInstallPackage(self, package, url, name, version, path):
        return 461, (package, ), self.stepInstallPackage, package, url, name, version, path

    def _getSuccess(self, success):
        return success and len(self._aborted) == 0

    def _getHeader(self):
        code = 471 if len(self._aborted) else 472
        return self.resolver.resolveString(code)

    def _getResult(self):
        modules = self._aborted if len(self._aborted) else self._installed
        return ', '.join(modules.keys())

