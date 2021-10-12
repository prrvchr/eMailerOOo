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

from .configuration import g_dns
from .configuration import g_extension
from .configuration import g_fetchsize
from .configuration import g_identifier
from .configuration import g_ispdb_page
from .configuration import g_ispdb_paths
from .configuration import g_merger_page
from .configuration import g_merger_paths

from .logger import Logger
from .logger import LogHandler
from .logger import Pool

from .logger import clearLogger
from .logger import getLoggerSetting
from .logger import getLoggerUrl
from .logger import getMessage
from .logger import isDebugMode
from .logger import logMessage
from .logger import setDebugMode
from .logger import setLoggerSetting

from .unotool import createService
from .unotool import executeDispatch
from .unotool import executeFrameDispatch
from .unotool import executeShell
from .unotool import getConfiguration
from .unotool import getConnectionMode
from .unotool import getContainerWindow
from .unotool import getDesktop
from .unotool import getDialog
from .unotool import getExceptionMessage
from .unotool import getFileSequence
from .unotool import getFileUrl
from .unotool import getInteractionHandler
from .unotool import getPathSettings
from .unotool import getPropertyValue
from .unotool import getPropertyValueSet
from .unotool import getResourceLocation
from .unotool import getSimpleFile
from .unotool import getStringResource
from .unotool import getUrl
from .unotool import getUrlPresentation
from .unotool import getUrlTransformer
from .unotool import hasInterface
from .unotool import parseUrl

from .mailerlib import Authenticator
from .mailerlib import CurrentContext
from .mailerlib import MailTransferable

from .mailertool import getDocument
from .mailertool import getDocumentFilter
from .mailertool import getMail
from .mailertool import getNamedExtension
from .mailertool import getUrlMimeType
from .mailertool import saveDocumentAs
from .mailertool import saveDocumentTmp
from .mailertool import saveTempDocument

from .oauth2tool import getOAuth2

from .dbtool import getObjectSequenceFromResult
from .dbtool import getSequenceFromResult
from .dbtool import getValueFromResult
from .dbtool import getTableColumns
from .dbtool import getTablesInfos

from .dbqueries import getSqlQuery

from .unolib import KeyMap

from .grid import GridManager
from .grid import GridListener
from .grid import RowSetListener

from .mail import MailModel
from .mail import MailManager
from .mail import MailView

from .datasource import DataSource
from .dataparser import DataParser

from .wizard import Wizard

from .smtpdispatch import SmtpDispatch

from .listener import TerminateListener

from .mailspooler import MailSpooler

from .ispdb import IspdbModel

from . import smtplib
