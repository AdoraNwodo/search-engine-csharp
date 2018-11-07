/// <summary>******************************************************************
/// 
/// Copyright (C) 2005  Stefano Franco
///
/// Based on JExcelAPI by Andrew Khan.
/// 
/// This library is free software; you can redistribute it and/or
/// modify it under the terms of the GNU Lesser General Public
/// License as published by the Free Software Foundation; either
/// version 2.1 of the License, or (at your option) any later version.
/// 
/// This library is distributed in the hope that it will be useful,
/// but WITHOUT ANY WARRANTY; without even the implied warranty of
/// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
/// Lesser General Public License for more details.
/// 
/// You should have received a copy of the GNU Lesser General Public
/// License along with this library; if not, write to the Free Software
/// Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
/// *************************************************************************
/// </summary>
using System;
using System.Collections;
using common;
using NExcel;
namespace NExcel.Biff.Formula
{
	
	/// <summary> An enumeration detailing the Excel function codes</summary>
	class Function
	{
		/// <summary> Gets the function code - used when generating token data
		/// 
		/// </summary>
		/// <returns> the code
		/// </returns>
		virtual internal int Code
		{
			get
			{
				return code;
			}
			
		}
		/// <summary> Gets the property name. Used by the FunctionNames object when initializing
		/// the locale specific names
		/// 
		/// </summary>
		/// <returns> the property name for this function
		/// </returns>
		virtual internal string PropertyName
		{
			get
			{
				return name;
			}
			
		}
		/// <summary> Gets the number of arguments for this function</summary>
		virtual internal int NumArgs
		{
			get
			{
				return numArgs;
			}
			
		}
		/// <summary> The logger</summary>
				private static Logger logger;
		
		/// <summary> The code which applies to this function</summary>
				private int code;
		
		/// <summary> The property name of this function</summary>
				private string name;
		
		/// <summary> The number of args this function expects</summary>
				private int numArgs;
		
		
		/// <summary> All available functions.  This attribute is package protected in order
		/// to enable the FunctionNames to initialize
		/// </summary>
		internal static Function[] functions;
		
		
		/// <summary> Constructor
		/// Sets the token value and adds this token to the array of all token
		/// 
		/// </summary>
		/// <param name="v">the biff code for the token
		/// </param>
		private Function(int v, string s, int a)
		{
			code = v;
			name = s;
			numArgs = a;
			
			// Grow the array
			Function[] newarray = new Function[functions.Length + 1];
			Array.Copy(functions, 0, newarray, 0, functions.Length);
			newarray[functions.Length] = this;
			functions = newarray;
		}
		
		/// <summary> Standard hash code method
		/// 
		/// </summary>
		/// <returns> the hash code
		/// </returns>
		public override int GetHashCode()
		{
			return code;
		}
		
		/// <summary> Gets the function name</summary>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <returns> the function name
		/// </returns>
		internal virtual string getName(WorkbookSettings ws)
		{
			FunctionNames fn = ws.FunctionNames;
			return fn.getName(this);
		}
		
		/// <summary> Gets the type object from its integer value</summary>
		public static Function getFunction(int v)
		{
			Function f = null;
			
			for (int i = 0; i < functions.Length; i++)
			{
				if (functions[i].code == v)
				{
					f = functions[i];
					break;
				}
			}
			
			return f != null?f:UNKNOWN;
		}
		
		/// <summary> Gets the type object from its string value.  Used when parsing strings</summary>
		/// <param name="">v
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <returns> the function 
		/// </returns>
		public static Function getFunction(string v, WorkbookSettings ws)
		{
			FunctionNames fn = ws.FunctionNames;
			Function f = fn.getFunction(v);
			return f != null?f:UNKNOWN;
		}
		
		// The functions
		
				public static readonly Function COUNT;
				public static readonly Function ATTRIBUTE;
				public static readonly Function ISNA;
				public static readonly Function ISERROR;
				public static readonly Function SUM;
				public static readonly Function AVERAGE;
				public static readonly Function MIN;
				public static readonly Function MAX;
				public static readonly Function ROW;
				public static readonly Function COLUMN;
				public static readonly Function NA;
				public static readonly Function NPV;
				public static readonly Function STDEV;
				public static readonly Function DOLLAR;
				public static readonly Function FIXED;
				public static readonly Function SIN;
				public static readonly Function COS;
				public static readonly Function TAN;
				public static readonly Function ATAN;
				public static readonly Function PI;
				public static readonly Function SQRT;
				public static readonly Function EXP;
				public static readonly Function LN;
				public static readonly Function LOG10;
				public static readonly Function ABS;
				public static readonly Function INT;
				public static readonly Function SIGN;
				public static readonly Function ROUND;
		//public static final Function LOOKUP;
				public static readonly Function INDEX;
		//public static final Function REPT;
				public static readonly Function MID;
				public static readonly Function LEN;
				public static readonly Function VALUE;
				public static readonly Function TRUE_Renamed;
				public static readonly Function FALSE_Renamed;
				public static readonly Function AND;
				public static readonly Function OR;
				public static readonly Function NOT;
				public static readonly Function MOD;
				public static readonly Function DCOUNT;
				public static readonly Function DSUM;
				public static readonly Function DAVERAGE;
				public static readonly Function DMIN;
				public static readonly Function DMAX;
				public static readonly Function DSTDEV;
				public static readonly Function VAR;
				public static readonly Function DVAR;
				public static readonly Function TEXT;
				public static readonly Function LINEST;
				public static readonly Function TREND;
				public static readonly Function LOGEST;
				public static readonly Function GROWTH;
		//public static final Function GOTO;
		//public static final Function HALT;
				public static readonly Function PV;
				public static readonly Function FV;
				public static readonly Function NPER;
				public static readonly Function PMT;
				public static readonly Function RATE;
		//public static final Function MIRR;
		//public static final Function IRR;
				public static readonly Function RAND;
				public static readonly Function MATCH;
				public static readonly Function DATE;
				public static readonly Function TIME;
				public static readonly Function DAY;
				public static readonly Function MONTH;
				public static readonly Function YEAR;
				public static readonly Function WEEKDAY;
				public static readonly Function HOUR;
				public static readonly Function MINUTE;
				public static readonly Function SECOND;
				public static readonly Function NOW;
				public static readonly Function AREAS;
				public static readonly Function ROWS;
				public static readonly Function COLUMNS;
				public static readonly Function OFFSET;
		//public static final Function ABSREF;
		//public static final Function RELREF;
		//public static final Function ARGUMENT;
		//public static final Function SEARCH;
				public static readonly Function TRANSPOSE;
				public static readonly Function ERROR;
		//public static final Function STEP;
				public static readonly Function TYPE;
		//public static final Function ECHO;
		//public static final Function SETNAME;
		//public static final Function CALLER;
		//public static final Function DEREF;
		//public static final Function WINDOWS;
		//public static final Function SERIES;
		//public static final Function DOCUMENTS;
		//public static final Function ACTIVECELL;
		//public static final Function SELECTION;
		//public static final Function RESULT;
				public static readonly Function ATAN2;
				public static readonly Function ASIN;
				public static readonly Function ACOS;
				public static readonly Function CHOOSE;
				public static readonly Function HLOOKUP;
				public static readonly Function VLOOKUP;
		//public static final Function LINKS;
		//public static final Function INPUT;
				public static readonly Function ISREF;
		//public static final Function GETFORMULA;
		//public static final Function GETNAME;
		//public static final Function SETVALUE;
				public static readonly Function LOG;
		//public static final Function EXEC;
				public static readonly Function CHAR;
				public static readonly Function LOWER;
				public static readonly Function UPPER;
				public static readonly Function PROPER;
				public static readonly Function LEFT;
				public static readonly Function RIGHT;
				public static readonly Function EXACT;
				public static readonly Function TRIM;
				public static readonly Function REPLACE;
				public static readonly Function SUBSTITUTE;
				public static readonly Function CODE;
		//public static final Function NAMES;
		//public static final Function DIRECTORY;
				public static readonly Function FIND;
				public static readonly Function CELL;
				public static readonly Function ISERR;
				public static readonly Function ISTEXT;
				public static readonly Function ISNUMBER;
				public static readonly Function ISBLANK;
				public static readonly Function T;
				public static readonly Function N;
		//public static final Function FOPEN;
		//public static final Function FCLOSE;
		//public static final Function FSIZE;
		//public static final Function FREADLN;
		//public static final Function FREAD;
		//public static final Function FWRITELN;
		//public static final Function FWRITE;
		//public static final Function FPOS;
				public static readonly Function DATEVALUE;
				public static readonly Function TIMEVALUE;
				public static readonly Function SLN;
				public static readonly Function SYD;
				public static readonly Function DDB;
		//public static final Function GETDEF;
		//public static final Function REFTEXT;
		//public static final Function TEXTREF;
				public static readonly Function INDIRECT;
		//public static final Function REGISTER;
		//public static final Function CALL;
		//public static final Function ADDBAR;
		//public static final Function ADDMENU;
		//public static final Function ADDCOMMAND;
		//public static final Function ENABLECOMMAND;
		//public static final Function CHECKCOMMAND;
		//public static final Function RENAMECOMMAND;
		//public static final Function SHOWBAR;
		//public static final Function DELETEMENU;
		//public static final Function DELETECOMMAND;
		//public static final Function GETCHARTITEM;
		//public static final Function DIALOGBOX;
				public static readonly Function CLEAN;
				public static readonly Function MDETERM;
				public static readonly Function MINVERSE;
				public static readonly Function MMULT;
		//public static final Function FILES;
				public static readonly Function IPMT;
				public static readonly Function PPMT;
				public static readonly Function COUNTA;
				public static readonly Function PRODUCT;
				public static readonly Function FACT;
		//public static final Function GETCELL;
		//public static final Function GETWORKSPACE;
		//public static final Function GETWINDOW;
		//public static final Function GETDOCUMENT;
				public static readonly Function DPRODUCT;
				public static readonly Function ISNONTEXT;
		//public static final Function GETNOTE;
		//public static final Function NOTE;
				public static readonly Function STDEVP;
				public static readonly Function VARP;
				public static readonly Function DSTDEVP;
				public static readonly Function DVARP;
				public static readonly Function TRUNC;
				public static readonly Function ISLOGICAL;
				public static readonly Function DCOUNTA;
				public static readonly Function FINDB;
				public static readonly Function SEARCHB;
				public static readonly Function REPLACEB;
				public static readonly Function LEFTB;
				public static readonly Function RIGHTB;
				public static readonly Function MIDB;
				public static readonly Function LENB;
				public static readonly Function ROUNDUP;
				public static readonly Function ROUNDDOWN;
				public static readonly Function RANK;
				public static readonly Function ADDRESS;
				public static readonly Function AYS360;
				public static readonly Function ODAY;
				public static readonly Function VDB;
				public static readonly Function MEDIAN;
				public static readonly Function SUMPRODUCT;
				public static readonly Function SINH;
				public static readonly Function COSH;
				public static readonly Function TANH;
				public static readonly Function ASINH;
				public static readonly Function ACOSH;
				public static readonly Function ATANH;
				public static readonly Function AVEDEV;
				public static readonly Function BETADIST;
				public static readonly Function GAMMALN;
				public static readonly Function BETAINV;
				public static readonly Function BINOMDIST;
				public static readonly Function CHIDIST;
				public static readonly Function CHIINV;
				public static readonly Function COMBIN;
				public static readonly Function CONFIDENCE;
				public static readonly Function CRITBINOM;
				public static readonly Function EVEN;
				public static readonly Function EXPONDIST;
				public static readonly Function FDIST;
				public static readonly Function FINV;
				public static readonly Function FISHER;
				public static readonly Function FISHERINV;
				public static readonly Function FLOOR;
				public static readonly Function GAMMADIST;
				public static readonly Function GAMMAINV;
				public static readonly Function CEILING;
				public static readonly Function HYPGEOMDIST;
				public static readonly Function LOGNORMDIST;
				public static readonly Function LOGINV;
				public static readonly Function NEGBINOMDIST;
				public static readonly Function NORMDIST;
				public static readonly Function NORMSDIST;
				public static readonly Function NORMINV;
				public static readonly Function NORMSINV;
				public static readonly Function STANDARDIZE;
				public static readonly Function ODD;
				public static readonly Function PERMUT;
				public static readonly Function POISSON;
				public static readonly Function TDIST;
				public static readonly Function WEIBULL;
				public static readonly Function SUMXMY2;
				public static readonly Function SUMX2MY2;
				public static readonly Function SUMX2PY2;
				public static readonly Function CHITEST;
				public static readonly Function CORREL;
				public static readonly Function COVAR;
				public static readonly Function FORECAST;
				public static readonly Function FTEST;
				public static readonly Function INTERCEPT;
				public static readonly Function PEARSON;
				public static readonly Function RSQ;
				public static readonly Function STEYX;
				public static readonly Function SLOPE;
				public static readonly Function TTEST;
				public static readonly Function PROB;
				public static readonly Function DEVSQ;
				public static readonly Function GEOMEAN;
				public static readonly Function HARMEAN;
				public static readonly Function SUMSQ;
				public static readonly Function KURT;
				public static readonly Function SKEW;
				public static readonly Function ZTEST;
				public static readonly Function LARGE;
				public static readonly Function SMALL;
				public static readonly Function QUARTILE;
				public static readonly Function PERCENTILE;
				public static readonly Function PERCENTRANK;
				public static readonly Function MODE;
				public static readonly Function TRIMMEAN;
				public static readonly Function TINV;
				public static readonly Function CONCATENATE;
				public static readonly Function POWER;
				public static readonly Function RADIANS;
				public static readonly Function DEGREES;
				public static readonly Function SUBTOTAL;
				public static readonly Function SUMIF;
				public static readonly Function COUNTIF;
				public static readonly Function COUNTBLANK;
				public static readonly Function HYPERLINK;
				public static readonly Function AVERAGEA;
				public static readonly Function MAXA;
				public static readonly Function MINA;
				public static readonly Function STDEVPA;
				public static readonly Function VARPA;
				public static readonly Function STDEVA;
				public static readonly Function VARA;
		
		// If token.  This is not an excel assigned number, but one made up
		// in order that the if command may be recognized
				public static readonly Function IF;
		
		// Unknown token
				public static readonly Function UNKNOWN;

		static Function()
		{
			logger = Logger.getLogger(typeof(Function));
			functions = new Function[0];

				 COUNT = new Function(0x0, "count", 0xff);
				 ATTRIBUTE = new Function(0x1, "", 0xff);
				 ISNA = new Function(0x2, "isna", 1);
				 ISERROR = new Function(0x3, "iserror", 1);
				 SUM = new Function(0x4, "sum", 0xff);
				 AVERAGE = new Function(0x5, "average", 0xff);
				 MIN = new Function(0x6, "min", 0xff);
				 MAX = new Function(0x7, "max", 0xff);
				 ROW = new Function(0x8, "row", 0xff);
				 COLUMN = new Function(0x9, "column", 1);
				 NA = new Function(0xa, "na", 0);
				 NPV = new Function(0xb, "npv", 0xff);
				 STDEV = new Function(0xc, "stdev", 0xff);
				 DOLLAR = new Function(0xd, "dollar", 2);
				 FIXED = new Function(0xe, "fixed", 0xff);
				 SIN = new Function(0xf, "sin", 1);
				 COS = new Function(0x10, "cos", 1);
				 TAN = new Function(0x11, "tan", 1);
				 ATAN = new Function(0x12, "atan", 1);
				 PI = new Function(0x13, "pi", 0);
				 SQRT = new Function(0x14, "sqrt", 1);
				 EXP = new Function(0x15, "exp", 1);
				 LN = new Function(0x16, "ln", 1);
				 LOG10 = new Function(0x17, "log10", 1);
				 ABS = new Function(0x18, "abs", 1);
				 INT = new Function(0x19, "int", 1);
				 SIGN = new Function(0x1a, "sign", 1);
				 ROUND = new Function(0x1b, "round", 2);
		//public static final Function LOOKUP =  new Function(0x1c, "LOOKUP",);
				 INDEX = new Function(0x1d, "index", 3);
		//public static final Function REPT =  new Function(0x1e, "REPT",);
				 MID = new Function(0x1f, "mid", 3);
				 LEN = new Function(0x20, "len", 1);
				 VALUE = new Function(0x21, "value", 1);
				 TRUE_Renamed = new Function(0x22, "true", 0);
				 FALSE_Renamed = new Function(0x23, "false", 0);
				 AND = new Function(0x24, "and", 0xff);
				 OR = new Function(0x25, "or", 0xff);
				 NOT = new Function(0x26, "not", 1);
				 MOD = new Function(0x27, "mod", 2);
				 DCOUNT = new Function(0x28, "dcount", 3);
				 DSUM = new Function(0x29, "dsum", 3);
				 DAVERAGE = new Function(0x2a, "daverage", 3);
				 DMIN = new Function(0x2b, "dmin", 3);
				 DMAX = new Function(0x2c, "dmax", 3);
				 DSTDEV = new Function(0x2d, "dstdev", 3);
				 VAR = new Function(0x2e, "var", 0xff);
				 DVAR = new Function(0x2f, "dvar", 3);
				 TEXT = new Function(0x30, "text", 2);
				 LINEST = new Function(0x31, "linest", 0xff);
				 TREND = new Function(0x32, "trend", 0xff);
				 LOGEST = new Function(0x33, "logest", 0xff);
				 GROWTH = new Function(0x34, "growth", 0xff);
		//public static final Function GOTO =  new Function(0x35, "GOTO",);
		//public static final Function HALT =  new Function(0x36, "HALT",);
				 PV = new Function(0x38, "pv", 0xff);
				 FV = new Function(0x39, "fv", 0xff);
				 NPER = new Function(0x3a, "nper", 0xff);
				 PMT = new Function(0x3b, "pmt", 0xff);
				 RATE = new Function(0x3c, "rate", 0xff);
		//public static final Function MIRR =  new Function(0x3d, "MIRR",);
		//public static final Function IRR =  new Function(0x3e, "IRR",);
				 RAND = new Function(0x3f, "rand", 0);
				 MATCH = new Function(0x40, "match", 3);
				 DATE = new Function(0x41, "date", 3);
				 TIME = new Function(0x42, "time", 3);
				 DAY = new Function(0x43, "day", 1);
				 MONTH = new Function(0x44, "month", 1);
				 YEAR = new Function(0x45, "year", 1);
				 WEEKDAY = new Function(0x46, "weekday", 2);
				 HOUR = new Function(0x47, "hour", 1);
				 MINUTE = new Function(0x48, "minute", 1);
				 SECOND = new Function(0x49, "second", 1);
				 NOW = new Function(0x4a, "now", 0);
				 AREAS = new Function(0x4b, "areas", 0xff);
				 ROWS = new Function(0x4c, "rows", 0xff);
				 COLUMNS = new Function(0x4d, "columns", 0xff);
				 OFFSET = new Function(0x4e, "offset", 0xff);
		//public static final Function ABSREF =  new Function(0x4f, "ABSREF",);
		//public static final Function RELREF =  new Function(0x50, "RELREF",);
		//public static final Function ARGUMENT =  new Function(0x51,"ARGUMENT",);
		//public static final Function SEARCH =  new Function(0x52, "SEARCH",3);
				 TRANSPOSE = new Function(0x53, "transpose", 0xff);
				 ERROR = new Function(0x54, "error", 1);
		//public static final Function STEP =  new Function(0x55, "STEP",);
				 TYPE = new Function(0x56, "type", 1);
		//public static final Function ECHO =  new Function(0x57, "ECHO",);
		//public static final Function SETNAME =  new Function(0x58, "SETNAME",);
		//public static final Function CALLER =  new Function(0x59, "CALLER",);
		//public static final Function DEREF =  new Function(0x5a, "DEREF",);
		//public static final Function WINDOWS =  new Function(0x5b, "WINDOWS",);
		//public static final Function SERIES =  new Function(0x5c, "SERIES",);
		//public static final Function DOCUMENTS =  new Function(0x5d,"DOCUMENTS",);
		//public static final Function ACTIVECELL =  new Function(0x5e,"ACTIVECELL",);
		//public static final Function SELECTION =  new Function(0x5f,"SELECTION",);
		//public static final Function RESULT =  new Function(0x60, "RESULT",);
				 ATAN2 = new Function(0x61, "atan2", 1);
				 ASIN = new Function(0x62, "asin", 1);
				 ACOS = new Function(0x63, "acos", 1);
				 CHOOSE = new Function(0x64, "choose", 0xff);
				 HLOOKUP = new Function(0x65, "hlookup", 0xff);
				 VLOOKUP = new Function(0x66, "vlookup", 0xff);
		//public static final Function LINKS =  new Function(0x67, "LINKS",);
		//public static final Function INPUT =  new Function(0x68, "INPUT",);
				 ISREF = new Function(0x69, "isref", 1);
		//public static final Function GETFORMULA =  new Function(0x6a,"GETFORMULA",);
		//public static final Function GETNAME =  new Function(0x6b, "GETNAME",);
		//public static final Function SETVALUE =  new Function(0x6c,"SETVALUE",);
				 LOG = new Function(0x6d, "log", 0xff);
		//public static final Function EXEC =  new Function(0x6e, "EXEC",);
				 CHAR = new Function(0x6f, "char", 1);
				 LOWER = new Function(0x70, "lower", 1);
				 UPPER = new Function(0x71, "upper", 1);
				 PROPER = new Function(0x72, "proper", 1);
				 LEFT = new Function(0x73, "left", 0xff);
				 RIGHT = new Function(0x74, "right", 0xff);
				 EXACT = new Function(0x75, "exact", 2);
				 TRIM = new Function(0x76, "trim", 1);
				 REPLACE = new Function(0x77, "replace", 4);
				 SUBSTITUTE = new Function(0x78, "substitute", 0xff);
				 CODE = new Function(0x79, "code", 1);
		//public static final Function NAMES =  new Function(0x7a, "NAMES",);
		//public static final Function DIRECTORY =  new Function(0x7b,"DIRECTORY",);
				 FIND = new Function(0x7c, "find", 0xff);
				 CELL = new Function(0x7d, "cell", 2);
				 ISERR = new Function(0x7e, "iserr", 1);
				 ISTEXT = new Function(0x7f, "istext", 1);
				 ISNUMBER = new Function(0x80, "isnumber", 1);
				 ISBLANK = new Function(0x81, "isblank", 1);
				 T = new Function(0x82, "t", 1);
				 N = new Function(0x83, "n", 1);
		//public static final Function FOPEN =  new Function(0x84, "FOPEN",);
		//public static final Function FCLOSE =  new Function(0x85, "FCLOSE",);
		//public static final Function FSIZE =  new Function(0x86, "FSIZE",);
		//public static final Function FREADLN =  new Function(0x87, "FREADLN",);
		//public static final Function FREAD =  new Function(0x88, "FREAD",);
		//public static final Function FWRITELN =  new Function(0x89,"FWRITELN",);
		//public static final Function FWRITE =  new Function(0x8a, "FWRITE",);
		//public static final Function FPOS =  new Function(0x8b, "FPOS",);
				 DATEVALUE = new Function(0x8c, "datevalue", 1);
				 TIMEVALUE = new Function(0x8d, "timevalue", 1);
				 SLN = new Function(0x8e, "sln", 3);
				 SYD = new Function(0x8f, "syd", 3);
				 DDB = new Function(0x90, "ddb", 0xff);
		//public static final Function GETDEF =  new Function(0x91, "GETDEF",);
		//public static final Function REFTEXT =  new Function(0x92, "REFTEXT",);
		//public static final Function TEXTREF =  new Function(0x93, "TEXTREF",);
				 INDIRECT = new Function(0x94, "indirect", 0xff);
		//public static final Function REGISTER =  new Function(0x95,"REGISTER",);
		//public static final Function CALL =  new Function(0x96, "CALL",);
		//public static final Function ADDBAR =  new Function(0x97, "ADDBAR",);
		//public static final Function ADDMENU =  new Function(0x98, "ADDMENU",);
		//public static final Function ADDCOMMAND =  new Function(0x99,"ADDCOMMAND",);
		//public static final Function ENABLECOMMAND =  new Function(0x9a,"ENABLECOMMAND",);
		//public static final Function CHECKCOMMAND =  new Function(0x9b,"CHECKCOMMAND",);
		//public static final Function RENAMECOMMAND =  new Function(0x9c,"RENAMECOMMAND",);
		//public static final Function SHOWBAR =  new Function(0x9d, "SHOWBAR",);
		//public static final Function DELETEMENU =  new Function(0x9e,"DELETEMENU",);
		//public static final Function DELETECOMMAND =  new Function(0x9f,"DELETECOMMAND",);
		//public static final Function GETCHARTITEM =  new Function(0xa0,"GETCHARTITEM",);
		//public static final Function DIALOGBOX =  new Function(0xa1,"DIALOGBOX",);
				 CLEAN = new Function(0xa2, "clean", 1);
				 MDETERM = new Function(0xa3, "mdeterm", 0xff);
				 MINVERSE = new Function(0xa4, "minverse", 0xff);
				 MMULT = new Function(0xa5, "mmult", 0xff);
		//public static final Function FILES =  new Function(0xa6, "FILES",);
				 IPMT = new Function(0xa7, "ipmt", 0xff);
				 PPMT = new Function(0xa8, "ppmt", 0xff);
				 COUNTA = new Function(0xa9, "counta", 0xff);
				 PRODUCT = new Function(0xb7, "product", 0xff);
				 FACT = new Function(0xb8, "fact", 1);
		//public static final Function GETCELL =  new Function(0xb9, "GETCELL",);
		//public static final Function GETWORKSPACE =  new Function(0xba,"GETWORKSPACE",);
		//public static final Function GETWINDOW =  new Function(0xbb,"GETWINDOW",);
		//public static final Function GETDOCUMENT =  new Function(0xbc,"GETDOCUMENT",);
				 DPRODUCT = new Function(0xbd, "dproduct", 3);
				 ISNONTEXT = new Function(0xbe, "isnontext", 1);
		//public static final Function GETNOTE =  new Function(0xbf, "GETNOTE",);
		//public static final Function NOTE =  new Function(0xc0, "NOTE",);
				 STDEVP = new Function(0xc1, "stdevp", 0xff);
				 VARP = new Function(0xc2, "varp", 0xff);
				 DSTDEVP = new Function(0xc3, "dstdevp", 0xff);
				 DVARP = new Function(0xc4, "dvarp", 0xff);
				 TRUNC = new Function(0xc5, "trunc", 0xff);
				 ISLOGICAL = new Function(0xc6, "islogical", 1);
				 DCOUNTA = new Function(0xc7, "dcounta", 0xff);
				 FINDB = new Function(0xcd, "findb", 0xff);
				 SEARCHB = new Function(0xce, "searchb", 3);
				 REPLACEB = new Function(0xcf, "replaceb", 4);
				 LEFTB = new Function(0xd0, "leftb", 0xff);
				 RIGHTB = new Function(0xd1, "rightb", 0xff);
				 MIDB = new Function(0xd2, "midb", 3);
				 LENB = new Function(0xd3, "lenb", 1);
				 ROUNDUP = new Function(0xd4, "roundup", 2);
				 ROUNDDOWN = new Function(0xd5, "rounddown", 2);
				 RANK = new Function(0xd8, "rank", 0xff);
				 ADDRESS = new Function(0xdb, "address", 0xff);
				 AYS360 = new Function(0xdc, "days360", 0xff);
				 ODAY = new Function(0xdd, "today", 0);
				 VDB = new Function(0xde, "vdb", 0xff);
				 MEDIAN = new Function(0xe3, "median", 0xff);
				 SUMPRODUCT = new Function(0xe4, "sumproduct", 0xff);
				 SINH = new Function(0xe5, "sinh", 1);
				 COSH = new Function(0xe6, "cosh", 1);
				 TANH = new Function(0xe7, "tanh", 1);
				 ASINH = new Function(0xe8, "asinh", 1);
				 ACOSH = new Function(0xe9, "acosh", 1);
				 ATANH = new Function(0xea, "atanh", 1);
				 AVEDEV = new Function(0x10d, "avedev", 0xFF);
				 BETADIST = new Function(0x10e, "betadist", 0xFF);
				 GAMMALN = new Function(0x10f, "gammaln", 1);
				 BETAINV = new Function(0x110, "betainv", 0xFF);
				 BINOMDIST = new Function(0x111, "binomdist", 4);
				 CHIDIST = new Function(0x112, "chidist", 2);
				 CHIINV = new Function(0x113, "chiinv", 2);
				 COMBIN = new Function(0x114, "combin", 2);
				 CONFIDENCE = new Function(0x115, "confidence", 3);
				 CRITBINOM = new Function(0x116, "critbinom", 3);
				 EVEN = new Function(0x117, "even", 1);
				 EXPONDIST = new Function(0x118, "expondist", 3);
				 FDIST = new Function(0x119, "fdist", 3);
				 FINV = new Function(0x11a, "finv", 3);
				 FISHER = new Function(0x11b, "fisher", 1);
				 FISHERINV = new Function(0x11c, "fisherinv", 1);
				 FLOOR = new Function(0x11d, "floor", 2);
				 GAMMADIST = new Function(0x11e, "gammadist", 4);
				 GAMMAINV = new Function(0x11f, "gammainv", 3);
				 CEILING = new Function(0x120, "ceiling", 2);
				 HYPGEOMDIST = new Function(0x121, "hypgeomdist", 4);
				 LOGNORMDIST = new Function(0x122, "lognormdist", 3);
				 LOGINV = new Function(0x123, "loginv", 3);
				 NEGBINOMDIST = new Function(0x124, "negbinomdist", 3);
				 NORMDIST = new Function(0x125, "normdist", 4);
				 NORMSDIST = new Function(0x126, "normsdist", 1);
				 NORMINV = new Function(0x127, "norminv", 3);
				 NORMSINV = new Function(0x128, "normsinv", 1);
				 STANDARDIZE = new Function(0x129, "standardize", 3);
				 ODD = new Function(0x12a, "odd", 1);
				 PERMUT = new Function(0x12b, "permut", 2);
				 POISSON = new Function(0x12c, "poisson", 3);
				 TDIST = new Function(0x12d, "tdist", 3);
				 WEIBULL = new Function(0x12e, "weibull", 4);
				 SUMXMY2 = new Function(303, "sumxmy2", 0xff);
				 SUMX2MY2 = new Function(304, "sumx2my2", 0xff);
				 SUMX2PY2 = new Function(305, "sumx2py2", 0xff);
				 CHITEST = new Function(0x132, "chitest", 0xff);
				 CORREL = new Function(0x133, "correl", 0xff);
				 COVAR = new Function(0x134, "covar", 0xff);
				 FORECAST = new Function(0x135, "forecast", 0xff);
				 FTEST = new Function(0x136, "ftest", 0xff);
				 INTERCEPT = new Function(0x137, "intercept", 0xff);
				 PEARSON = new Function(0x138, "pearson", 0xff);
				 RSQ = new Function(0x139, "rsq", 0xff);
				 STEYX = new Function(0x13a, "steyx", 0xff);
				 SLOPE = new Function(0x13b, "slope", 2);
				 TTEST = new Function(0x13c, "ttest", 0xff);
				 PROB = new Function(0x13d, "prob", 0xff);
				 DEVSQ = new Function(0x13e, "devsq", 0xff);
				 GEOMEAN = new Function(0x13f, "geomean", 0xff);
				 HARMEAN = new Function(0x140, "harmean", 0xff);
				 SUMSQ = new Function(0x141, "sumsq", 0xff);
				 KURT = new Function(0x142, "kurt", 0xff);
				 SKEW = new Function(0x143, "skew", 0xff);
				 ZTEST = new Function(0x144, "ztest", 0xff);
				 LARGE = new Function(0x145, "large", 0xff);
				 SMALL = new Function(0x146, "small", 0xff);
				 QUARTILE = new Function(0x147, "quartile", 0xff);
				 PERCENTILE = new Function(0x148, "percentile", 0xff);
				 PERCENTRANK = new Function(0x149, "percentrank", 0xff);
				 MODE = new Function(0x14a, "mode", 0xff);
				 TRIMMEAN = new Function(0x14b, "trimmean", 0xff);
				 TINV = new Function(0x14c, "tinv", 2);
				 CONCATENATE = new Function(0x150, "concatenate", 0xff);
				 POWER = new Function(0x151, "power", 2);
				 RADIANS = new Function(0x156, "radians", 1);
				 DEGREES = new Function(0x157, "degrees", 1);
				 SUBTOTAL = new Function(0x158, "subtotal", 0xff);
				 SUMIF = new Function(0x159, "sumif", 0xff);
				 COUNTIF = new Function(0x15a, "countif", 2);
				 COUNTBLANK = new Function(0x15b, "countblank", 0xff);
				 HYPERLINK = new Function(0x167, "hyperlink", 2);
				 AVERAGEA = new Function(0x169, "averagea", 0xff);
				 MAXA = new Function(0x16a, "maxa", 0xff);
				 MINA = new Function(0x16b, "mina", 0xff);
				 STDEVPA = new Function(0x16c, "stdevpa", 0xff);
				 VARPA = new Function(0x16d, "varpa", 0xff);
				 STDEVA = new Function(0x16e, "stdeva", 0xff);
				 VARA = new Function(0x16f, "vara", 0xff);

		// If token.  This is not an excel assigned number, but one made up
		// in order that the if command may be recognized
				 IF = new Function(0xfffe, "if", 0xff);
		
		// Unknown token
				 UNKNOWN = new Function(0xffff, "", 0);

		}
	}
}
