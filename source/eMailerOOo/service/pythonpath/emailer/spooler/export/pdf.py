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

from ...unotool import getConfiguration
from ...unotool import getPropertyValueSet

import traceback


class PdfExport():
    def __init__(self, ctx):
        entry = 'org.openoffice.Office.Common'
        path = 'Filter/PDF/Export'
        descriptor = self._getDescriptor(ctx, entry, path)
        self._descriptor = getPropertyValueSet(descriptor)

    def getDescriptor(self):
        return self._descriptor

    def _getDescriptor(self, ctx, entry, path):
        descriptor = {}
        config = self._getConfig(ctx, entry, path)
        if config:
            self._setProperties(config, descriptor, self._getGeneralProperties())
            self._setProperties(config, descriptor, self._getViewProperties())
            self._setProperties(config, descriptor, self._getInterfaceProperties())
            self._setProperties(config, descriptor, self._getLinkProperties())
            self._setProperties(config, descriptor, self._getSecurityProperties())
        return descriptor

    def _getConfig(self, ctx, entry, path):
        configuration = getConfiguration(ctx, entry, False)
        if configuration.hasByHierarchicalName(path):
            config = configuration.getByHierarchicalName(path)
        else:
            config = None
        return config

    def _setProperties(self, config, descriptor, properties):
        for property in properties:
            if config.hasByName(property):
                descriptor[property] = config.getByName(property)

    def _getGeneralProperties(self):
        return ('UseLosslessCompression',
                'Quality',
                'ReduceImageResolution',
                'MaxImageResolution',
                'SelectPdfVersion',
                'UseTaggedPDF',
                'ExportFormFields',
                'FormsType',
                'AllowDuplicateFieldNames',
                'ExportBookmarks',
                'ExportNotes',
                'ExportNotesPages',
                'IsSkipEmptyPages',
                'IsAddStream')

    def _getViewProperties(self):
        return ('InitialView',
                'InitialPage',
                'Magnification',
                'Zoom',
                'PageLayout',
                'FirstPageOnLeft')

    def _getInterfaceProperties(self):
        return ('ResizeWindowToInitialPage',
                'CenterWindow',
                'OpenInFullScreenMode',
                'DisplayPDFDocumentTitle',
                'HideViewerMenubar',
                'HideViewerToolbar',
                'HideViewerWindowControls',
                'UseTransitionEffects',
                'OpenBookmarkLevels')

    def _getLinkProperties(self):
        return ('ExportBookmarksToPDFDestination',
                'ConvertOOoTargetToPDFTarget',
                'ExportLinksRelativeFsys',
                'PDFViewSelection')

    def _getSecurityProperties(self):
        return ('Printing',
                'Changes',
                'EnableCopyingOfContent',
                'EnableTextAccessForAccessibilityTools')

