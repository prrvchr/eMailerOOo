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

from ..configuration import g_chunk

import ijson
import traceback


def parseMessage(response):
    messageid = None
    labels = []
    events = ijson.sendable_list()
    parser = ijson.parse_coro(events)
    iterator = response.iterContent(g_chunk, False)
    while iterator.hasMoreElements():
        chunk = iterator.nextElement().value
        print("ApiHelper.parseMessage() Content:\n%s" % chunk.decode('utf-8'))
        parser.send(chunk)
        for prefix, event, value in events:
            print("ApiHelper.parseMessage() Prefix: %s - Event: %s - Value: %s" % (prefix, event, value))
            if (prefix, event) == ('id', 'string'):
                messageid = value
            elif (prefix, event) == ('labelIds.item', 'string'):
                labels.append(value)
        del events[:]
    parser.close()
    return messageid, labels

def setDefaultFolder(server, url, messageid, labels, folder='SENT'):
    labelids = []
    for label in labels:
        if label != folder:
            labelids.append(label)
    if labelids:
        parameter = server.getRequestParameter('getDefaultFolder')
        parameter.Method = 'POST'
        parameter.Url = url + '%s/modify' % messageid
        parameter.Json = '{"addLabelIds": [], "removeLabelIds": ["%s"]}' % '","'.join(labelids)
        response = server.execute(parameter)

