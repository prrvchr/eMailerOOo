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

from ..unotool import deregisterStartupJob
from ..unotool import getPathSubstitution
from ..unotool import getResourceLocation
from ..unotool import getSimpleFile
from ..unotool import getStringResource

from ..logger import getLogger

from ..configuration import g_basename
from ..configuration import g_defaultlog
from ..configuration import g_identifier

from packaging.requirements import Requirement
from time import sleep
import importlib
import pkg_resources as pkgr
import os
import sys
import subprocess
import traceback


class SetupModel():
    def __init__(self, ctx, job, name, code):
        self._ctx = ctx
        self._job = job
        self._code = code
        self._modules = []
        self._requirements = '/requirements.txt'
        self._program = getPathSubstitution(ctx, '$(prog)')
        self._url = getResourceLocation(self._ctx, g_identifier)
        self._pythonpath = self._url + '/service/pythonpath'
        self._pip = self._pythonpath + '/pip/__main__.py'
        self._logger = getLogger(ctx, g_defaultlog, g_basename)
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', name)
        self._resources = {'Title': 'SetupWindow.Title',
                           'Header': 'SetupWindow.Label1.Label.%s',
                           'Text': 'SetupWindow.Label2.Label',
                           'Install': 'SetupWindow.Label4.Label'}

    def getPage(self, step):
        return step, self._getHeader(step)

    def checkRequirements(self, maxProgress, progress):
        self._modules = []
        url = self._url + self._requirements
        if getSimpleFile(self._ctx).exists(url):
            self._checkPackages(url, maxProgress, progress)
        success = len(self._modules) > 0
        return success, self._getResult(self._modules)

    def installPackages(self, maxProgress, progress):
        index = 1
        maxProgress(len(self._modules) + 1)
        progress(self._getInstallText(''), index)
        modules = []
        isWindows = os.name == 'nt'
        info = self._getStartupInfo(isWindows)
        command = self._getPipCommand(isWindows)
        for module in self._modules:
            index += 1
            progress(self._getInstallText(module), index)
            message = self._pipInstall(info, command, module)
            if message is None:
                importlib.invalidate_caches()
                self._log(INFO, self._code + 1, module)
            else:
                modules.append(module)
                self._log(SEVERE, self._code + 2, module, message)
        success = len(modules) == 0
        return success, self._getResult(self._modules) if success else self._getResult(modules)

    def deregisterJob(self):
        deregisterStartupJob(self._ctx, self._job)

    def _checkPackages(self, url, maxProgress, progress):
        packages = []
        with open(uno.fileUrlToSystemPath(url)) as requirements:
            for requirement in pkgr.parse_requirements(requirements):
                packages.append(requirement.project_name)
        index = 1
        maxProgress(len(packages))
        for package in packages:
            module = self._getInstallModule(package)
            progress(self._getProgessText(package, module), index)
            index += 1

    def _getInstallModule(self, package):
        requirement = Requirement(package)
        target = requirement.name.lower().replace('-', '_')
        for module, packages in importlib.metadata.packages_distributions().items():
            if target in [p.lower().replace('-', '_') for p in packages]:
                return module
        self._modules.append(target)
        return target

    def _getPipCommand(self, isWindows):
        if isWindows:
            executable = uno.fileUrlToSystemPath(self._program + '/python.exe')
        else:
            executable = uno.fileUrlToSystemPath(self._program + '/python')
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

    def _getResult(self, modules):
        return ', '.join(modules) if len(modules) else 'None'

    def _log(self, level, code, *args):
        self._logger.logprb(level, 'SetupModel', 'installModules()', code, *args)

# SetupModel StringRessoure methods
    def getTitle(self):
        return self._resolver.resolveString(self._resources.get('Title'))

    def _getHeader(self, step):
        return self._resolver.resolveString(self._resources.get('Header') % step)

    def _getProgessText(self, *args):
        return self._resolver.resolveString(self._resources.get('Text')) % args

    def _getInstallText(self, module):
        return self._resolver.resolveString(self._resources.get('Install')) + module
