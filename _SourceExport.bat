@echo off
cd %~dp0

rmdir "Source\Ladex.xlam"

cd ..
cscript //nologo vbac.wsf decombine /binary:Ladex /source:Ladex/Source



cd %~dp0
rmdir /s /q "Source\メンテナンス用.xlsm"
