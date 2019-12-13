
cd .\smtpMailerOOo

..\zip.exe -0 smtpMailerOOo.zip mimetype

..\zip.exe -r smtpMailerOOo.zip *

cd ..

move /Y .\smtpMailerOOo\smtpMailerOOo.zip .\smtpMailerOOo.oxt
