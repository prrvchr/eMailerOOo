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

from ..unotool import checkVersion
from ..unotool import createService
from ..unotool import deregisterStartupJob
from ..unotool import getConfiguration
from ..unotool import getExtensionVersion
from ..unotool import getPathSubstitution
from ..unotool import getResourceLocation
from ..unotool import getSimpleFile
from ..unotool import getStringResource
from ..unotool import saveWindowPosition

from ..logger import getLogger

from ..configuration import g_basename
from ..configuration import g_defaultlog
from ..configuration import g_identifier

from packaging.requirements import Requirement
import importlib
import pkg_resources as pkgr
from time import sleep
import os
import re
import sys
import subprocess
import traceback


class SetupModel():
    def __init__(self, ctx, name, code, java, database, dependencies):
        self._ctx = ctx
        self._name = name
        self._code = code
        self._java = java
        self._database = java and database
        self._dependencies = dependencies
        self._modules = []
        self._closing = False
        self._position = 'SetupPosition'
        self._requirements = '/requirements.txt'
        self._program = getPathSubstitution(ctx, '$(prog)')
        self._url = getResourceLocation(self._ctx, g_identifier)
        self._pythonpath = self._url + '/service/pythonpath'
        self._pip = self._pythonpath + '/pip/__main__.py'
        self._logger = getLogger(ctx, g_defaultlog, g_basename)
        self._config = getConfiguration(ctx, g_identifier, True)
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'SetupWindow')
        self._resources = {'Title': 'SetupWindow.Title',
                           'Header': 'SetupWindow.Label1.Label.%s',
                           'Dependency': 'SetupWindow.Label2.Label',
                           'Java': 'SetupWindow.Label5.Label',
                           'Text': 'SetupWindow.Label12.Label',
                           'Install': 'SetupWindow.Label15.Label'}
    def close(self):
        self._closing = True

    def getPage(self, step):
        return step, self._getHeader(step)

    def getViewData(self):
        return self._getDialogPosition(), self._getTitle()

    def savePosition(self, position):
        saveWindowPosition(self._config, position, self._position)

    def hasDataBase(self):
        return self._database

    def hasJava(self):
        return self._java

    def hasDependencies(self):
        return len(self._dependencies) != 0 

    def checkDependencies(self, maxProgress, progress):
        i = 0
        dependencies = []
        maxProgress(len(self._dependencies))
        for identifier, data in self._dependencies.items():
            if self._closing:
                break
            i += 1
            progress(self._getDependencyText(data), i)
            if not self._checkDependency(identifier, data):
                dependencies.append(data)
            sleep(1)
        success = len(dependencies) == 0
        return success, self._getDependenciesResult(self._dependencies.values() if success else dependencies)

    def checkJava(self, maxProgress, progress):
        from ..jdbcdriver import g_java
        success = False
        maxProgress(2)
        progress(self._getJavaText(), 1)
        version = self._getJavaVersion()
        progress(self._getJavaText(version), 2)
        sleep(1)
        success = version is not None and checkVersion(version, g_java)
        return success, self._getJavaResult(version if success else g_java)

    def _getJavaVersion(self):
        version = None
        try:
            resultat = subprocess.run(['java', '-version'],
                                      stdout=subprocess.PIPE,
                                      stderr=subprocess.PIPE,
                                      text=True,
                                      check=True)
            line = resultat.stderr.splitlines()[0]
            match = re.search(r'"([^"]+)"', line)
            if match:
                version = match.group(1)
        except Exception:
            pass
        return version

    def checkRequirements(self, maxProgress, progress):
        packages = []
        self._modules = []
        url = self._url + self._requirements
        if getSimpleFile(self._ctx).exists(url):
            self._checkPackages(packages, url, maxProgress, progress)
        success = len(self._modules) == 0
        return success, self._getRequirementsResult(packages if success else self._modules)

    def installPackages(self, maxProgress, progress):
        size = len(self._modules) * 2
        maxProgress(size)
        i = 0
        modules = []
        isWindows = os.name == 'nt'
        info = self._getStartupInfo(isWindows)
        command = self._getPipCommand(isWindows)
        for module in self._modules:
            if self._closing:
                break
            i += 1
            progress(self._getInstallText(module), i)
            message = self._pipInstall(info, command, module)
            i += 1
            progress(self._getInstallText(module), i)
            if message is None:
                importlib.invalidate_caches()
                self._log(INFO, self._code + 1, module)
            else:
                modules.append(module)
                self._log(SEVERE, self._code + 2, module, message)
            if i == size:
                sleep(1)
        success = len(modules) == 0
        return success, self._getRequirementsResult(self._modules if success else modules)

    def deregisterJob(self):
        deregisterStartupJob(self._ctx, self._name)

    def _checkDependency(self, identifier, data):
        version = getExtensionVersion(self._ctx, identifier)
        return version is not None and checkVersion(version, data[1])

    def _checkPackages(self, packages, url, maxProgress, progress):
        with open(uno.fileUrlToSystemPath(url)) as requirements:
            for requirement in pkgr.parse_requirements(requirements):
                if self._closing:
                    break
                packages.append(requirement.project_name)
        i = 0
        size = len(packages)
        maxProgress(size)
        for package in packages:
            if self._closing:
                break
            i += 1
            module = self._getInstallModule(package)
            progress(self._getProgessText(package, module), i)
            if i == size:
                sleep(1)

    def _getInstallModule(self, package):
        target = self._parseModuleName(Requirement(package).name)
        for module, packages in importlib.metadata.packages_distributions().items():
            if self._closing:
                break
            if target in [self._parseModuleName(p) for p in packages]:
                return module
        else:
            self._modules.append(target)
        return target

    def _parseModuleName(self, module):
        return module.lower().replace('-', '_')

    def _getPipCommand(self, isWindows):
        if isWindows:
            url = self._program + '/python.exe'
        else:
            url = self._program + '/python'
        executable = uno.fileUrlToSystemPath(url)
        pip = uno.fileUrlToSystemPath(self._pip)
        path = uno.fileUrlToSystemPath(self._pythonpath)
        return [executable, '-u', pip, 'install', '--target', path, '--only-binary=:all:', 'module']

    def _getStartupInfo(self, isWindows):
        info = None
        if isWindows:
            info = subprocess.STARTUPINFO()
            info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        return info

    def _pipInstall(self, info, command, module):
        msg = None
        try:
            command[-1] = module
            process = subprocess.Popen(command,
                                       stdout=subprocess.PIPE,
                                       stderr=subprocess.PIPE,
                                       startupinfo=info)
            stdout, stderr = process.communicate()
            if process.returncode != 0:
                msg = stderr.decode('utf-8', errors='ignore')
        except Exception as e:
            pass
        return msg

    def _getRequirementsResult(self, modules):
        return ', '.join(modules)

    def _getDependenciesResult(self, dependencies):
        return '\n'.join(['%s - version %s' % (name, version) for (name, version) in dependencies])

    def _getJavaResult(self, version):
        return 'Java JDK version: ' + version

    def _log(self, level, code, *args):
        self._logger.logprb(level, 'SetupModel', 'installModules()', code, *args)

    def _getDialogPosition(self):
        return uno.createUnoStruct('com.sun.star.awt.Point', *self._config.getByName(self._position))

# SetupModel StringRessoure methods
    def _getTitle(self):
        return self._resolver.resolveString(self._resources.get('Title'))

    def _getHeader(self, step):
        return self._resolver.resolveString(self._resources.get('Header') % step)

    def _getDependencyText(self, data):
        return self._resolver.resolveString(self._resources.get('Dependency')) % data

    def _getJavaText(self, version=''):
        return self._resolver.resolveString(self._resources.get('Java')) + version

    def _getProgessText(self, *args):
        return self._resolver.resolveString(self._resources.get('Text')) % args

    def _getInstallText(self, module):
        return self._resolver.resolveString(self._resources.get('Install')) + module

