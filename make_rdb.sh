#!/bin/bash

OOoProgram=/usr/lib/libreoffice/program
Path=$(dirname "${0}")
File=${Path}/rdb/types.rdb

rm -f ${File}

# ./rdb/make_rdb.sh com/sun/star/auth/XRestKeyMap
# ./rdb/make_rdb.sh com/sun/star/auth/OAuth2Request
# ./rdb/make_rdb.sh com/sun/star/auth/XInteractionUserName
# ./rdb/make_rdb.sh com/sun/star/auth/RestRequestTokenType

read -p "Press enter to continue"

if test -f "${File}"; then
    ${OOoProgram}/regview ${File}
fi
