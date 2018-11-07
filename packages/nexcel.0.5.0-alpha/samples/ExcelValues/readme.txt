NExcel  -  A Open-Source .NET library for reading Ms Excel files
Copyright (c) 2005  Stefano Franco

ExcelValues  -  sample application for reading Excel values


--------------------------------------------------------------------
INTRODUCTION

This folder contains ExcelValues, a sample application for reading values from Excel files using NExcel. 
ExcelValues reads a Excel file and writes to console all cell's values.

ExcelValues is a basic sample for using NExcel. It shows how using property "Cell.Value".
It can be useful to get started and for debugging purposes.

To execute it, open a DOS console, go to this folder and execute the command 

  ExcelValues  sample.xls


--------------------------------------------------------------------
USAGE

Usage: ExcelValues  filename

  filename             the Excel file name


Example:

  ExcelValues  sample.xls


--------------------------------------------------------------------
SOURCE 

Source is available with a ready-to-use VS.NET 2003 project.

Otherwise in this folder is also available a build batch file, "build.bat". 
To build with it, open file build.bat, set its program's file paths (see "constants")
and then execute it.

