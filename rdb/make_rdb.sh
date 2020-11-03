#!/bin/bash

OOoPath=/usr/lib/libreoffice
Path=$(dirname "${0}")

rm -f ${Path}/types.rdb

${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XMailService2
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XSmtpService2
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XMailServiceProvider2

read -p "Press enter to continue"

if test -f "${Path}/types.rdb"; then
    ${OOoPath}/program/regview ${Path}/types.rdb
    read -p "Press enter to continue"
fi
