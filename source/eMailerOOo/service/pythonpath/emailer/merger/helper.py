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

from com.sun.star.awt import Point

from .mergercontroller import MergerController

from ..wizard import Wizard

from ..unotool import getConfiguration
from ..unotool import getDesktop
from ..unotool import showMessageBox

from ..configuration import g_basename
from ..configuration import g_identifier
from ..configuration import g_mergerframe
from ..configuration import g_merger_page
from ..configuration import g_merger_paths
from ..configuration import g_resource

import traceback


def showMergerWizard(ctx, **_):
    try:
        frame = getDesktop(ctx).getCurrentFrame()
        if frame:
            document = frame.getController().getModel()
            if document.hasLocation():
                point = Point(*getConfiguration(ctx, g_identifier).getByName('MergerPosition'))
                wizard = Wizard(ctx, g_merger_page, name=g_mergerframe, point=point)
                controller = MergerController(ctx, wizard, document)
                arguments = (g_merger_paths, controller)
                wizard.initialize(arguments)
                wizard.execute()
            else:
                showMessageBox(ctx, 1221, g_identifier, g_resource, g_basename, title=document.Title)
        else:
            showMessageBox(ctx, 1211, g_identifier, g_resource, g_basename)
    except Exception:
        print("helper.showMergerWizard() ERROR: %s" % traceback.format_exc())
