::@echo off
set nowPath=%cd%
cd \
cd %nowPath%

echo %nowPath%

for /r %nowPath% %%i in (*.pdb,*.vshost.*) do (del %%i) 

for /r %nowPath% %%i in (obj,bin,.vs,DTAR_08E86330_4835_4B5C_9E5A_61F37AE1A077_DTAR,Lib\NPOI\solution\Lib) do (IF EXIST %%i RD /s /q %%i) 

echo OK
