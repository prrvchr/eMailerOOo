#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020 https://prrvchr.github.io                                     ║
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

# General configuration
g_extension = 'eMailerOOo'
g_identifier = 'io.github.prrvchr.%s' % g_extension
# Ispdb roadmap wizard paths (Online and Offline)
g_ispdb_paths = ((1, 2, 3, 4, 5), (1, 2, 3, 4))
g_ispdb_page = -1
# Merger roadmap wizard paths
g_merger_paths = (1, 2, 3)
g_merger_page = -1
# Grid RowSet.FetchSize
g_fetchsize = 500
# Internet DNS connection
g_dns = ('1.1.1.1', 53)
# Thread Logo
g_logo = 'eMailerOOo.png'
g_logourl = 'https://prrvchr.github.io/%s/img/%s' %(g_extension, g_logo)
# Logger resource strings files folder
g_resource = 'resource'
# Logger resource strings files basename
g_basename = 'eMailer'
# Logger configuration
g_defaultlog = 'eMailerLogger'
g_errorlog = 'eMailerError'
g_spoolerlog = 'SpoolerLogger'
g_mailservicelog = 'MailServiceLogger'
g_chunk = 320 * 1024
