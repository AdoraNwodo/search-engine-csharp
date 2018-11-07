@echo off

rem ---------------------------------------------------
rem   This file builds ExcelContents.exe
rem ---------------------------------------------------
 

rem  Constants
rem  The dir of csc.exe
set CSCBIN=%windir%\Microsoft.NET\Framework\v1.0.3705\csc.exe


rem building ExcelContents.exe
echo Building ExcelContents.exe ....
"%CSCBIN%" /nologo /out:ExcelContents.exe  /t:exe  /r:NExcel.dll  /recurse:*.cs
echo done.
