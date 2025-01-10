#!
# -*- coding: utf-8 -*-

"""
╔════════════════════════════════════════════════════════════════════════════════════╗
║                                                                                    ║
║   Copyright (c) 2020-24 https://prrvchr.github.io                                  ║
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

from com.sun.star.uno import Exception as UnoException

from com.sun.star.mail import MailException

from ..unotool import getConfiguration
from ..unotool import getNamedValueSet
from ..unotool import hasInterface

from ..mailerlib import CustomMessage

from ..oauth2 import CustomParser

from ..oauth2 import getParserItems
from ..oauth2 import getRequest
from ..oauth2 import getResponseResults
from ..oauth2 import setResquestParameter

from ..configuration import g_chunk
from ..configuration import g_identifier

from string import Template
import ijson
import traceback


def getHttpProvider(providers, name):
    provider = url = None
    # FIXME: If the connection context does not have
    # FIXME: a Provider value then name can be null.
    if name and providers.hasByName(name):
        provider = name
        url = providers.getByName(name).getByName('Url')
    return provider, url

def getHttpServer(ctx, url, servername, username):
    server = url if url else servername
    return getRequest(ctx, server, username)

def getHttpRequest(provider, server, request):
    if provider.hasByName(server.value):
        service = provider.getByName(server.value)
        if service.hasByName(request):
            return service.getByName(request)
    return None

def executeHttpRequest(source, logger, debug, cls, code, server, message, request):
    mtd = 'executeHttpRequest'
    name = request.getByName('Name')
    parameter = server.getRequestParameter(name)
    arguments = CustomMessage(logger, cls, message, request)
    setResquestParameter(arguments, request, parameter)
    try:
        response = server.execute(parameter)
        response.raiseForStatus()
    except UnoException as e:
        msg = logger.resolveString(code, message.Subject, e.Message)
        if debug:
            logger.logp(SEVERE, cls, mtd, msg)
        raise MailException(msg, source)
    if response.Ok:
        parser = CustomParser(*getParserItems(request))
        # XXX: It may be possible that there is nothing to parse
        if parser.hasItems():
            results = getResponseResults(parser, response)
            interface = 'com.sun.star.mail.XMailMessage2'
            if hasInterface(message, interface):
                for name, value in results.items():
                    logger.logprb(INFO, cls, mtd, code + 1, name, value)
                    message.setHeader(name, value)
    response.close()

def getOAuth2TokenWithParameters(oauth2, parameters, port):
    return oauth2.getTokenWithParameters(parameters, getNamedValueSet({'Port': port}))

