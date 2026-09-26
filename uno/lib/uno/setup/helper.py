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

from ..unotool import checkVersion
from ..unotool import createService
from ..unotool import executeDesktopDispatch
from ..unotool import getExtensionVersion
from ..unotool import getPropertyValueSet
from ..unotool import getResourceLocation
from ..unotool import getSimpleFile

import importlib
import os
from packaging.requirements import Requirement
import pkg_resources as pkgr
import subprocess
import sys
from time import sleep
import traceback


def showSetup(ctx, identifier, listener=None, /, **kwargs):
    url = f'vnd.sun.star.job:service={identifier}.Setup'
    executeDesktopDispatch(ctx, url, listener, **kwargs)

def checkExtension(ctx, identifier, data):
    version = getExtensionVersion(ctx, identifier)
    return version is not None and checkVersion(version, data[1])

def getJavaStatus(ctx):
    service = 'com.sun.star.comp.stoc.JavaVirtualMachine'
    jvm = createService(ctx, service)
    if jvm is None:
        return 4
    if jvm.isVMEnabled():
        return 0
    return 3

def getJavaVersion(ctx, extension, java, script):
    results = 2, java, java
    try:
        service = 'com.sun.star.script.provider.MasterScriptProviderFactory'
        factory = createService(ctx, service)
        provider = factory.createScriptProvider('')
        url = f'vnd.sun.star.script:{script}?language=Java&location=user:uno_packages/{extension}.oxt'
        macro = provider.getScript(url)
        if macro:
            result = macro.invoke((), (), ())
            version = result[0]
            if version:
                if checkVersion(version, java):
                    results = 0, java, version
                else:
                    results = 1, java, version
    except Exception:
        pass
    return results

def checkAgent(ctx, service, url, agent):
    support = False
    driver = createService(ctx, service)
    if driver:
        properties = getPropertyValueSet({agent: True})
        for info in driver.getPropertyInfo(url, properties):
            if info.Name == agent:
                support = info.Value != 'false'
                break
    return support

def getPackageCount(url):
    i = 0
    with open(uno.fileUrlToSystemPath(url)) as requirements:
        for requirement in pkgr.parse_requirements(requirements):
            i += 1
    return i

def checkPython(url, addInstalled, addMissing):
    with open(uno.fileUrlToSystemPath(url)) as requirements:
        for requirement in pkgr.parse_requirements(requirements):
            target = _parseModuleName(Requirement(requirement.project_name).name)
            for module, packages in importlib.metadata.packages_distributions().items():
                if target in [_parseModuleName(p) for p in packages]:
                    yield addInstalled(module)
                    break
            else:
                yield addMissing(target)

def getPipCommand(program, isWindows, python):
    if isWindows:
        url = program + '/python.exe'
    else:
        url = program + '/python'
    executable = uno.fileUrlToSystemPath(url)
    if not os.path.exists(executable):
        executable = sys.executable
    pip = uno.fileUrlToSystemPath(python + '/pip/__main__.py')
    path = uno.fileUrlToSystemPath(python)
    return [executable, '-u', pip, 'install', '--target', path, '--only-binary=:all:', 'module']

def getStartupInfo(isWindows):
    info = None
    if isWindows:
        info = subprocess.STARTUPINFO()
        info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    return info

def pipInstall(info, command, module):
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
        print("helper.pipInstall() ERROR: %s" % traceback.format_exc())
        pass
    return msg

def _parseModuleName(module):
    return module.lower().replace('-', '_')

