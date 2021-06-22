@echo off
cd %~dp0

cd Sorce
del /Q "Microsoft Excel Objects"\*.*
del /Q Modules\*.*
del /Q Class\*.*
del /Q Form\*.*

cd ..\..
cscript //nologo vbac.wsf decombine /binary:Ladex /source:Ladex/Sorce

cd "Ladex\Sorce\Ladex.xlam"

move *.dcm "..\Microsoft Excel Objects"
move *.bas ..\Modules
move *.frm ..\Form
move *.frx ..\Form
move *.cls ..\Class


cd %~dp0
rmdir "Sorce\Ladex.xlam"
rmdir /s /q "Sorce\メンテナンス用.xlsm"
