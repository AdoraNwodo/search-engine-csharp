@echo off

rem ---------------------------------------------------
rem   This file builds ExcelValues.exe
rem ---------------------------------------------------
 

rem  Constants
rem  The dir of csc.exe
set CSCBIN=%windir%\Microsoft.NET\Framework\v1.0.3705\csc.exe


rem building ExcelValues.exe
echo Building ExcelValues.exe ....
"%CSCBIN%" /nologo /out:ExcelValues.exe  /t:exe  /r:NExcel.dll  /recurse:*.cs
echo done.
