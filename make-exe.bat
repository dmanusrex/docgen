@echo off

setlocal EnableDelayedExpansion

del docgen.exe

::: Build and Sign the exe
python build.py

::: Signing needs to be more dynamic...
signtool sign /a /s MY /n "Open Source Developer, Darren Richer" /fd SHA256 /t http://time.certum.pl /v docgen.exe

::: Clean up build artifacts
rmdir /q/s build

::: Create zip file
del docgen.zip
powershell Compress-Archive docgen.exe docgen.zip

::: Sign release artifact
::: del docgen.zip.asc
::: gpg --detach-sign --armor --local-user EB0F2232 swon-analyzer.zip
