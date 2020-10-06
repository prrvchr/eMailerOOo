#!/bin/bash

./rdb/make_rdb.sh

cd ./smtpMailerOOo/
zip -0 smtpMailerOOo.zip mimetype
zip -r smtpMailerOOo.zip *
cd ..

mv ./smtpMailerOOo/smtpMailerOOo.zip ./smtpMailerOOo.oxt
