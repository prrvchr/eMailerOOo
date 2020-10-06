#!/bin/bash

./rdb/make_rdb.sh

cd ./smtpServerOOo/
zip -0 smtpServerOOo.zip mimetype
zip -r smtpServerOOo.zip *
cd ..

mv ./smtpServerOOo/smtpServerOOo.zip ./smtpServerOOo.oxt
