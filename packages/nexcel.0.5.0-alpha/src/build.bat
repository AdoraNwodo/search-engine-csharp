@echo off

rem Builds NExcel.dll

rem  Constants
set RESGENBIN=C:\Program Files\Microsoft.NET\FrameworkSDK\Bin\ResGen.exe
set CSCBIN=%windir%\Microsoft.NET\Framework\v1.0.3705\csc.exe
set RESDIR=.\NExcel\Biff\Formula


echo  Building NExcel.dll ....

rem Building resources
rem converts .resx in .resources
copy  "%RESDIR%\FunctionNames.resx" .
"%RESGENBIN%"  FunctionNames.resx
del  FunctionNames.resx
ren FunctionNames.resources  NExcel.Biff.Formula.FunctionNames.resources
copy  "%RESDIR%\FunctionNames_fr.resx" .
"%RESGENBIN%"  FunctionNames_fr.resx
del  FunctionNames_fr.resx
ren FunctionNames_fr.resources  NExcel.Biff.Formula.FunctionNames_fr.resources
copy  "%RESDIR%\FunctionNames_es.resx" .
"%RESGENBIN%"  FunctionNames_es.resx
del  FunctionNames_es.resx
ren FunctionNames_es.resources  NExcel.Biff.Formula.FunctionNames_es.resources
copy  "%RESDIR%\FunctionNames_de.resx" .
"%RESGENBIN%"  FunctionNames_de.resx
del  FunctionNames_de.resx
ren FunctionNames_de.resources  NExcel.Biff.Formula.FunctionNames_de.resources

 
rem building NExcel.dll
"%CSCBIN%"  /out:NExcel.dll  /t:library /res:NExcel.Biff.Formula.FunctionNames.resources /res:NExcel.Biff.Formula.FunctionNames_fr.resources /res:NExcel.Biff.Formula.FunctionNames_es.resources /res:NExcel.Biff.Formula.FunctionNames_de.resources  /recurse:*.cs


rem removes temp files
del  NExcel.Biff.Formula.FunctionNames.resources
del  NExcel.Biff.Formula.FunctionNames_fr.resources
del  NExcel.Biff.Formula.FunctionNames_es.resources
del  NExcel.Biff.Formula.FunctionNames_de.resources

echo  done.
