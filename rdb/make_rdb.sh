#!/bin/bash
<<COMMENT
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
COMMENT
OOoPath=/usr/lib/libreoffice
Path=$(dirname "${0}")

rm -f ${Path}/types.rdb

${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XMailService2
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XSmtpService2
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XMailServiceProvider2
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XSpoolerListener
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XSpoolerService
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/MailServiceProvider2
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XIspDBService
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/IspDBService
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/SpoolerService
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/auth/OAuth2Request
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/auth/XInteractionUserName
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/auth/XRestDataParser
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/auth/XRestKeyMap

read -p "Press enter to continue"

if test -f "${Path}/types.rdb"; then
    ${OOoPath}/program/regview ${Path}/types.rdb
    read -p "Press enter to continue"
fi
