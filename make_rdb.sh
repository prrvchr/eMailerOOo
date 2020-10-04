#!/bin/bash

OOoProgram=/usr/lib/libreoffice/program
Path=$(dirname "${0}")

rm ${Path}/rdb/types.rdb

# ./rdb/make_rdb.sh com/sun/star/auth/XRestKeyMap
# ./rdb/make_rdb.sh com/sun/star/auth/OAuth2Request
# ./rdb/make_rdb.sh com/sun/star/auth/XInteractionUserName
# ./rdb/make_rdb.sh com/sun/star/auth/RestRequestTokenType

read -p "Press enter to continue"

${OOoProgram}/regview ${Path}/rdb/types.rdb
