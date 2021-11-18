@echo off
cd %~dp0

cd Source
del /Q "Microsoft Excel Objects"\*.*
del /Q Modules\*.*
del /Q Class\*.*
del /Q Form\*.*

cd ..\..
cscript //nologo vbac.wsf decombine /binary:Ladex /source:Ladex/Source

cd "Ladex\Source\Ladex.xlam"

move *.dcm "..\Microsoft Excel Objects"
move *.bas ..\Modules
move *.frm ..\Form
move *.frx ..\Form
move *.cls ..\Class


cd %~dp0
rmdir "Source\Ladex.xlam"
rmdir /s /q "Source\メンテナンス用.xlsm"
