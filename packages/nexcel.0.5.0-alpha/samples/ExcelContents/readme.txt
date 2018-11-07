NExcel  -  A Open-Source .NET library for reading Ms Excel files
Copyright (c) 2005  Stefano Franco

ExcelContents  -  sample application for reading Excel contents


--------------------------------------------------------------------
INTRODUCTION

This folder contains ExcelContents, a sample application for reading contents from Excel files using NExcel. 
ExcelContents reads a Excel file and writes to console all cell's contents. 

ExcelContents is a basic sample for using NExcel. It shows how using property "Cell.Contents". 
It can be useful to get started and for debugging purposes.

To execute it, open a DOS console, go to this folder and execute the command 

  ExcelContents  sample.xls


--------------------------------------------------------------------
USAGE

Usage: ExcelContents  filename

  filename             the Excel file name


Example:

  ExcelContents  sample.xls


--------------------------------------------------------------------
SOURCE 

Source is available with a ready-to-use VS.NET 2003 project.

Otherwise in this folder is also available a build batch file, "build.bat". 
To build with it, open file build.bat, set its program's file paths (see "constants")
and then execute it.


