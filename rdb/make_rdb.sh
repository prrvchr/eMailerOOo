#!/bin/bash

OOoPath=/usr/lib/libreoffice
Path=$(dirname "${0}")

rm -f ${Path}/types.rdb

${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XMailService2
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XSmtpService2
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XMailServiceProvider2
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/MailServiceProvider2
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/XIspDBService
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/mail/IspDBService
${Path}/merge_rdb.sh ${OOoPath} com/sun/star/auth/XRestKeyMap

read -p "Press enter to continue"

if test -f "${Path}/types.rdb"; then
    ${OOoPath}/program/regview ${Path}/types.rdb
    read -p "Press enter to continue"
fi
