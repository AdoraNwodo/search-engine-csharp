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
using System.Globalization;
using System.Collections;
using common;
using NExcel.Biff.Formula;
namespace NExcel
{
	
	/// <summary> This is a bean which client applications may use to set various advanced
	/// workbook properties.  Use of this bean is not mandatory, and its absence
	/// will merely result in workbooks being read/written using the default
	/// settings
	/// </summary>
	public sealed class WorkbookSettings
	{
		/// <summary> Accessor for the array grow size property
		/// 
		/// </summary>
		/// <returns> the array grow size
		/// </returns>
		/// <summary> Sets the amount of memory by which to increase the amount of
		/// memory allocated to storing the workbook data.
		/// For processeses reading many small workbooks
		/// inside  a WAS it might be necessary to reduce the default size
		/// Default value is 1 megabyte
		/// 
		/// </summary>
		/// <param name="sz">the file size in bytes
		/// </param>
		public int ArrayGrowSize
		{
			get
			{
				return arrayGrowSize;
			}
			
			set
			{
				arrayGrowSize = value;
			}
			
		}
		/// <summary> Accessor for the initial file size property
		/// 
		/// </summary>
		/// <returns> the initial file size
		/// </returns>
		/// <summary> Sets the initial amount of memory allocated to store the workbook data
		/// when reading a worksheet.  For processeses reading many small workbooks
		/// inside  a WAS it might be necessary to reduce the default size
		/// Default value is 5 megabytes
		/// 
		/// </summary>
		/// <param name="sz">the file size in bytes
		/// </param>
		public int InitialFileSize
		{
			get
			{
				return initialFileSize;
			}
			
			set
			{
				initialFileSize = value;
			}
			
		}
		/// <summary> Gets the drawings disabled flag
		/// 
		/// </summary>
		/// <returns> TRUE if drawings are disabled, FALSE otherwise
		/// </returns>
		public bool DrawingsDisabled
		{
			get
			{
				return drawingsDisabled;
			}
			
		}
		/// <summary> Accessor for the disabling of garbage collection
		/// 
		/// </summary>
		/// <returns> FALSE if JExcelApi hints for garbage collection, TRUE otherwise
		/// </returns>
		public bool GCDisabled
		{
			get
			{
				return gcDisabled;
			}
			
		}
		/// <summary> Accessor for the disabling of interpretation of named ranges
		/// 
		/// </summary>
		/// <returns> FALSE if named cells are interpreted, TRUE otherwise
		/// </returns>
		/// <summary> Disables the handling of names
		/// 
		/// </summary>
		/// <param name="b">TRUE to disable the names feature, FALSE otherwise
		/// </param>
		public bool NamesDisabled
		{
			get
			{
				return namesDisabled;
			}
			
			set
			{
				namesDisabled = value;
			}
			
		}
		/// <summary> Sets whether or not to rationalize the cell formats before
		/// writing out the sheet.  The default value is true
		/// 
		/// </summary>
		/// <param name="r">the rationalization flag
		/// </param>
		public bool Rationalization
		{
			set
			{
				rationalizationDisabled = !value;
			}
			
		}
		/// <summary> Accessor to retrieve the rationalization flag
		/// 
		/// </summary>
		/// <returns> TRUE if rationalization is off, FALSE if rationalization is on
		/// </returns>
		public bool RationalizationDisabled
		{
			get
			{
				return rationalizationDisabled;
			}
			
		}
		/// <summary> Accessor to set the suppress warnings flag.  Due to the change
		/// in logging in version 2.4, this will now set the warning
		/// behaviour across the JVM (depending on the type of logger used)
		/// 
		/// </summary>
		/// <param name="w">the flag
		/// </param>
		public bool SuppressWarnings
		{
			set
			{
				logger.SuppressWarnings = value;
			}
			
		}
		/// <summary> Accessor for the formula adjust disabled
		/// 
		/// </summary>
		/// <returns> TRUE if formulas are adjusted following row/column inserts/deletes
		/// FALSE otherwise
		/// </returns>
		/// <summary> Setter for the formula adjust disabled property
		/// 
		/// </summary>
		/// <param name="b">TRUE to adjust formulas, FALSE otherwise
		/// </param>
		public bool FormulaAdjust
		{
			get
			{
				return !formulaReferenceAdjustDisabled;
			}
			
			set
			{
				formulaReferenceAdjustDisabled = !value;
			}
			
		}
		/// <summary> Returns the locale used for the spreadsheet
		/// 
		/// </summary>
		/// <returns> the locale
		/// </returns>
		/// <summary> Sets the locale for this spreadsheet
		/// 
		/// </summary>
		/// <param name="l">the locale
		/// </param>
		public System.Globalization.CultureInfo Locale
		{
			get
			{
				return locale;
			}
			
			set
			{
				locale = value;
			}
			
		}
		/// <summary> Accessor for the character encoding
		/// 
		/// </summary>
		/// <returns> the character encoding for this workbook
		/// </returns>
		/// <summary> Sets the encoding for this workbook
		/// 
		/// </summary>
		/// <param name="enc">the encoding
		/// </param>
		public string Encoding
		{
			get
			{
				return encoding;
			}
			
			set
			{
				encoding = value;
			}
			
		}
		/// <summary> Gets the function names.  This is used by the formula parsing package
		/// in order to get the locale specific function names for this particular
		/// workbook
		/// 
		/// </summary>
		/// <returns> the list of function names
		/// </returns>
		public FunctionNames FunctionNames
		{
			get
			{
				if (functionNames == null)
				{
					functionNames = (FunctionNames) localeFunctionNames[locale];
					
					// have not previously accessed function names for this locale,
					// so create a brand new one and add it to the list
					if (functionNames == null)
					{
						functionNames = new FunctionNames(locale);
						localeFunctionNames[locale] =  functionNames;
					}
				}
				
				return functionNames;
			}
			
		}
		/// <summary> Accessor for the character set.   This value is only used for reading
		/// and has no effect when writing out the spreadsheet
		/// 
		/// </summary>
		/// <returns> the character set used by this spreadsheet
		/// </returns>
		/// <summary> Sets the character set.  This is only used when the spreadsheet is
		/// read, and has no effect when the spreadsheet is written
		/// 
		/// </summary>
		/// <param name="cs">the character set encoding value
		/// </param>
		public int CharacterSet
		{
			get
			{
				return characterSet;
			}
			
			set
			{
				characterSet = value;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The amount of memory allocated to store the workbook data when
		/// reading a worksheet.  For processeses reading many small workbooks inside
		/// a WAS it might be necessary to reduce the default size
		/// </summary>
		private int initialFileSize;
		
		/// <summary> The amount of memory allocated to the array containing the workbook
		/// data when its current amount is exhausted.
		/// </summary>
		private int arrayGrowSize;
		
		/// <summary> Flag to indicate whether the drawing feature is enabled or not
		/// Drawings deactivated using -Djxl.nodrawings=true on the JVM command line
		/// Activated by default or by using -Djxl.nodrawings=false on the JVM command
		/// line
		/// </summary>
		private bool drawingsDisabled;
		
		/// <summary> Flag to indicate whether the name feature is enabled or not
		/// Names deactivated using -Djxl.nonames=true on the JVM command line
		/// Activated by default or by using -Djxl.nonames=false on the JVM command
		/// line
		/// </summary>
		private bool namesDisabled;
		
		/// <summary> Flag to indicate whether formula cell references should be adjusted
		/// following row/column insertion/deletion
		/// </summary>
		private bool formulaReferenceAdjustDisabled;
		
		/// <summary> Flag to indicate whether the system hint garbage collection
		/// is enabled or not.
		/// As a rule of thumb, it is desirable to enable garbage collection
		/// when reading large spreadsheets from  a batch process or from the
		/// command line, but better to deactivate the feature when reading
		/// large spreadsheets within a WAS, as the calls to System.gc() not
		/// only garbage collect the junk in JExcelApi, but also in the
		/// webservers JVM and can cause significant slowdown
		/// GC deactivated using -Djxl.nogc=true on the JVM command line
		/// Activated by default or by using -Djxl.nogc=false on the JVM command line
		/// </summary>
		private bool gcDisabled;
		
		/// <summary> Flag to indicate whether the rationalization of cell formats is
		/// disabled or not.
		/// Rationalization is enabled by default, but may be disabled for
		/// performance reasons.  It can be deactivated using -Djxl.norat=true on
		/// the JVM command line
		/// </summary>
		private bool rationalizationDisabled;
		
		/// <summary> The locale.  Normally this is the same as the system locale, but there
		/// may be cases (eg. where you are uploading many spreadsheets from foreign
		/// sources) where you may want to specify the locale on an individual
		/// worksheet basis
		/// The locale may also be specified on the command line using the lang and
		/// country System properties eg. -Djxl.lang=en -Djxl.country=UK for UK
		/// English
		/// </summary>
		private System.Globalization.CultureInfo locale;
		
		/// <summary> The locale specific function names for this workbook</summary>
		private FunctionNames functionNames;
		
		/// <summary> The character encoding used for reading non-unicode strings.  This can
		/// be different from the default platform encoding if processing spreadsheets
		/// from abroad.  This may also be set using the system property NExcel.encoding
		/// </summary>
		private string encoding;
		
		/// <summary> The character set used by the readable spreadsheeet</summary>
		private int characterSet;
		
		/// <summary> A hash map of function names keyed on locale</summary>
		private Hashtable localeFunctionNames;
		
		// **
		// The default values
		// **
		private const int defaultInitialFileSize = 5 * 1024 * 1024;
		// 5 megabytes
		private const int defaultArrayGrowSize = 1024 * 1024; // 1 megabyte
		
		/// <summary> Default constructor</summary>
		public WorkbookSettings()
		{
			initialFileSize = defaultInitialFileSize;
			arrayGrowSize = defaultArrayGrowSize;
			localeFunctionNames = new Hashtable();
			
			// Initialize other properties from the system properties
			// [TODO-NExcel_Next] make it in C#
//			try
//			{
//				      boolean suppressWarnings = Boolean.getBoolean("NExcel.nowarnings");
//				      setSuppressWarnings(suppressWarnings);
//				      drawingsDisabled        = Boolean.getBoolean("NExcel.nodrawings");
//				      namesDisabled           = Boolean.getBoolean("NExcel.nonames");
//				      gcDisabled              = Boolean.getBoolean("NExcel.nogc");
//				      rationalizationDisabled = Boolean.getBoolean("NExcel.norat");
//				      formulaReferenceAdjustDisabled =
//				                                Boolean.getBoolean("NExcel.noformulaadjust");
//				
//				      encoding = System.getProperty("file.encoding");
//			}
//			catch (System.Security.SecurityException e)
//			{
//				logger.warn("Error accessing system properties.", e);
//			}
			
			// Initialize the locale to the system locale
			try
			{
				// [TODO-NExcel_Next] make it in C#
				//	  if (System.getProperty("NExcel.lang")    == null ||
				//          System.getProperty("NExcel.country") == null)
				//      {
				//        locale = Locale.getDefault();
				//      }
				//      else
				//      {
				//        locale = new Locale(System.getProperty("NExcel.lang"),
				//                            System.getProperty("NExcel.country"));
				//      }
				//
				//      if (System.getProperty("NExcel.encoding") != null)
				//      {
				//        encoding = System.getProperty("NExcel.encoding");
				//      }
				// [TODO] check it - it's critical
				// the temporary old one
//				locale = new Globalization.CultureInfo("en-US"); 
				locale = CultureInfo.CurrentUICulture;
				// the default
				this.encoding = "ascii";

			}
			catch (System.Security.SecurityException e)
			{
				logger.warn("Error accessing system properties.", e);
				// [TODO] check it - it's critical
				//      locale = Locale.getDefault();
				// the temporary old one
//				locale = new System.Globalization.CultureInfo("en-US"); 
				locale = CultureInfo.InvariantCulture;
				this.encoding = "ascii";
			}
		}
		static WorkbookSettings()
		{
			logger = Logger.getLogger(typeof(WorkbookSettings));
		}
	}
}
