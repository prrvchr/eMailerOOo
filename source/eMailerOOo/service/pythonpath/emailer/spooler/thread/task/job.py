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

from ....unotool import getLastNamedParts

import traceback


class Job():
    def __init__(self, sf, url, folder, merge, filter):
        self._url = url
        tmp, name = getLastNamedParts(url, '/')
        self._name = name
        temp = None
        if merge or filter:
            temp = self._getTempFolder(folder, name)
            if not sf.exists(temp):
                sf.createFolder(temp)
        self._folder = temp
        self._merge = merge
        self._filter = filter

    @property
    def Url(self):
        return self._url

    @Url.setter
    def Url(self, url):
        self._url = url

    @property
    def Name(self):
        return self._name

    @Name.setter
    def Name(self, name):
        self._name = name

    @property
    def Folder(self):
        return self._folder

    @property
    def Merge(self):
        return self._merge

    @property
    def Filter(self):
        return self._filter

    def getUrl(self, job):
        return self._url

# Private shared methods
    def _getTempFolder(self, folder, name):
        return folder

