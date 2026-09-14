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

from ..unotool import getFileUrl
from ..unotool import getPathSettings
from ..unotool import getStringResource

from ..configuration import g_identifier

from collections import OrderedDict
import json

import traceback


class FilePicker():
    def __init__(self, ctx, path=None, **kwargs):
        self._ctx = ctx
        if path is None:
            path = getPathSettings(ctx).Work
        self._path = path
        self._resolver = getStringResource(ctx, g_identifier, 'dialogs', 'MessageBox')
        self._resources = {'PickerTitle':     'Mail.FilePicker.Title',
                           'PickerFilters':   'Mail.FilePicker.Filters'}

    def getDocumentUrl(self):
        title = self._getFilePickerTitle()
        filters = self._getFilePickerFilters()
        url, self._path = getFileUrl(self._ctx, title, self._path, filters)
        return url

    def getPath(self):
        return self._path

    def _getFilePickerTitle(self):
        resource = self._resources.get('PickerTitle')
        title = self._resolver.resolveString(resource)
        return title

    def _getFilePickerFilters(self):
        resource = self._resources.get('PickerFilters')
        filter = self._resolver.resolveString(resource)
        filters = json.loads(filter, object_pairs_hook=OrderedDict)
        return filters.items()

