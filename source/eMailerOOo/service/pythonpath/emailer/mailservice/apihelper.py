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

from ..unotool import getConfiguration

from ..mailertool import setParametersArguments

from ..mailerlib import CustomMessage

from ..oauth2 import getRequest

from ..configuration import g_chunk
from ..configuration import g_identifier

from string import Template
import ijson
import traceback


def getHttpProvider(providers, servername):
    for provider in providers:
        if servername in providers[provider]:
            return provider
    return None

def getHttpRequest(ctx, provider, servername, username):
    config = getConfiguration(ctx, g_identifier)
    url = config.getByName('Providers').getByName(provider).getByName('Url')
    server = url if url else servername
    return getRequest(ctx, server, username)

def getHttpRequests(ctx, provider, server):
    config = getConfiguration(ctx, g_identifier)
    if config.getByName('Providers').getByName(provider).hasByName('Services'):
        services = config.getByName('Providers').getByName(provider).getByName('Services')
        if services and services.hasByName(server.value):
            service = services.getByName(server.value)
            if service and service.hasByName('Requests'):
                return service.getByName('Requests')
    return None

def setResquestParameter(logger, cls, request, parameter, message):
    print("apihelper.setResquestParameter()")
    mtd = 'setResquestParameter'
    arguments = CustomMessage(logger, cls, message)
    setParametersArguments(request.getByName('Parameters'), arguments)
    method = request.getByName('Method')
    if method:
        parameter.Method = method
    url = request.getByName('Url')
    if url:
        parameter.Url = Template(url).safe_substitute(arguments)
    data = request.getByName('Data')
    if data and data in arguments:
        parameter.Data = uno.ByteSequence(arguments[data])
    template = request.getByName('Arguments')
    if template:
        config = arguments.toJson(template)
        parameter.fromJson(config)

def getResponseResults(items, response):
    results = {}
    events = ijson.sendable_list()
    parser = ijson.parse_coro(events)
    iterator = response.iterContent(g_chunk, False)
    while iterator.hasMoreElements():
        chunk = iterator.nextElement().value
        print("apihelper.getResponseResults() 1 chunk: %s" % chunk.decode())
        parser.send(chunk)
        for prefix, event, value in events:
            print("apihelper.getResponseResults() 2 prefix: %s - event: %s - value: %s" % (prefix, event, value))
            items.parse(results, prefix, event, value)
        del events[:]
    parser.close()
    return results

def getParserItems(request):
    keys = {}
    items = {}
    triggers = {}
    responses = request.getByName('Responses')
    if responses:
        for name in responses.getElementNames():
            response = responses.getByName(name)
            item = response.getByName('Item')
            if item and len(item) == 2:
                trigger = response.getByName('Trigger')
                if not trigger:
                    items[item] = name
                elif len(trigger) == 3:
                    keys[name] = item
                    triggers[trigger] = name
    return keys, items, triggers

