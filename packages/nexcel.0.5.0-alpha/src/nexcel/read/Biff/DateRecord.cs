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
using NExcelUtils;
using common;
using NExcel;
using NExcel.Format;
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	/// <summary> A date which is stored in the cell</summary>
	class DateRecord : DateCell
	{
		/// <summary> Interface method which returns the row number of this cell
		/// 
		/// </summary>
		/// <returns> the zero base row number
		/// </returns>
		virtual public int Row
		{
			get
			{
				return row;
			}
			
		}
		/// <summary> Interface method which returns the column number of this cell
		/// 
		/// </summary>
		/// <returns> the zero based column number
		/// </returns>
		virtual public int Column
		{
			get
			{
				return column;
			}
			
		}
		/// <summary> Gets the date
		/// 
		/// </summary>
		/// <returns> the date
		/// </returns>
		virtual public DateTime DateValue
		{
			get
			{
				return _Value;
			}
			
		}
		/// <summary> Gets the cell contents as a string.  This method will use the java
		/// equivalent of the excel formatting string
		/// 
		/// </summary>
		/// <returns> the label
		/// </returns>
		virtual public string Contents
		{
			get
			{
				return string.Format(format, "{0}", _Value);
			}
			
		}

		/// <summary>
		/// Returns a date.
		/// </summary>
		virtual public object Value
		{
			get
			{
				return this._Value;
			}
		}


		/// <summary> Accessor for the cell type
		/// 
		/// </summary>
		/// <returns> the cell type
		/// </returns>
		virtual public CellType Type
		{
			get
			{
				return CellType.DATE;
			}
			
		}
		/// <summary> Indicates whether the date value contained in this cell refers to a date,
		/// or merely a time
		/// 
		/// </summary>
		/// <returns> TRUE if the value refers to a time
		/// </returns>
		virtual public bool Time
		{
			get
			{
				return time;
			}
			
		}
		/// <summary> Gets the DateFormat used to format the cell.  This will normally be
		/// the format specified in the excel spreadsheet, but in the event of any
		/// difficulty parsing this, it will revert to the default date/time format.
		/// 
		/// </summary>
		/// <returns> the DateFormat object used to format the date in the original
		/// excel cell
		/// </returns>
		virtual public DateTimeFormatInfo DateFormat
		{
			get
			{
				Assert.verify(format != null);
				
				return format;
			}
			
		}
		/// <summary> Gets the CellFormat object for this cell.  Used by the WritableWorkbook
		/// API
		/// 
		/// </summary>
		/// <returns> the CellFormat used for this cell
		/// </returns>
		virtual public NExcel.Format.CellFormat CellFormat
		{
			get
			{
				if (!initialized)
				{
					cellFormat = formattingRecords.getXFRecord(xfIndex);
					initialized = true;
				}
				
				return cellFormat;
			}
			
		}
		/// <summary> Determines whether or not this cell has been hidden
		/// 
		/// </summary>
		/// <returns> TRUE if this cell has been hidden, FALSE otherwise
		/// </returns>
		virtual public bool Hidden
		{
			get
			{
				ColumnInfoRecord cir = sheet.getColumnInfo(column);
				
				if (cir != null && cir.Width == 0)
				{
					return true;
				}
				
				RowRecord rr = sheet.getRowInfo(row);
				
				if (rr != null && (rr.RowHeight == 0 || rr.isCollapsed()))
				{
					return true;
				}
				
				return false;
			}
			
		}
		/// <summary> Accessor for the sheet
		/// 
		/// </summary>
		/// <returns>  the containing sheet
		/// </returns>
		virtual protected internal SheetImpl Sheet
		{
			get
			{
				return sheet;
			}
			
		}
		/// <summary> The date represented within this cell</summary>
		private DateTime _Value;
		/// <summary> The row number of this cell record</summary>
		private int row;
		/// <summary> The column number of this cell record</summary>
		private int column;
		
		/// <summary> Indicates whether this is a full date, or merely a time</summary>
		private bool time;
		
		/// <summary> The format to use when displaying this cell's contents as a string</summary>
		private DateTimeFormatInfo format;
		
		/// <summary> The raw cell format</summary>
		private NExcel.Format.CellFormat cellFormat;
		
		/// <summary> The index to the XF Record</summary>
		private int xfIndex;
		
		/// <summary> A handle to the formatting records</summary>
		private FormattingRecords formattingRecords;
		
		/// <summary> A handle to the sheet</summary>
		private SheetImpl sheet;
		
		/// <summary> A flag to indicate whether this objects formatting things have
		/// been initialized
		/// </summary>
		private bool initialized;
		
		// The default formats used when returning the date as a string
		private static readonly DateTimeFormatInfo dateFormat = new DateTimeFormatInfo("dd MMM yyyy");
		private static readonly DateTimeFormatInfo timeFormat = new DateTimeFormatInfo("HH:mm:ss");
		
		// The number of days between 1 Jan 1900 and 1 March 1900. Excel thinks
		// the day before this was 29th Feb 1900, but it was 28th Feb 1900.
		// I guess the programmers thought nobody would notice that they
		// couldn't be bothered to program this dating anomaly properly
		private const int nonLeapDay = 61;
		
		// [TODO-NExcel_Next]
//		private static readonly System.TimeZone gmtZone;
		
		// The number of days between 01 Jan 1900 and 01 Jan 1970 - this gives
		// the UTC offset
		private const int utcOffsetDays = 25569;
		
		// The number of days between 01 Jan 1904 and 01 Jan 1970 - this gives
		// the UTC offset using the 1904 date system
		private const int utcOffsetDays1904 = 24107;
		
		// The number of milliseconds in  a day
		private const long msInADay = 24 * 60 * 60 * 1000;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="num">the numerical representation of this
		/// </param>
		/// <param name="xfi">the java equivalent of the excel date format
		/// </param>
		/// <param name="fr"> the formatting records
		/// </param>
		/// <param name="nf"> flag indicating whether we are using the 1904 date system
		/// </param>
		/// <param name="si"> the sheet
		/// </param>
		public DateRecord(NumberCell num, int xfi, FormattingRecords fr, bool nf, SheetImpl si)
		{
			row = num.Row;
			column = num.Column;
			xfIndex = xfi;
			formattingRecords = fr;
			sheet = si;
			initialized = false;
			
			format = formattingRecords.getDateFormat(xfIndex);
			
			// This value represents the number of days since 01 Jan 1900
			double numValue = num.DoubleValue;
			
			// Work round a bug in excel.  Excel seems to think there is a date
			// called the 29th Feb, 1900 - but in actual fact this was not a leap year.
			// Therefore for values less than 61 in the 1900 date system,
			// add one to the numeric value
			if (!nf && numValue < nonLeapDay)
			{
				numValue += 1;
			}
			
			if (System.Math.Abs(numValue) < 1)
			{
				if (format == null)
				{
					format = timeFormat;
				}
				time = true;
			}
			else
			{
				if (format == null)
				{
					format = dateFormat;
				}
				time = false;
			}
			
			// Get rid of any timezone adjustments - we are not interested
			// in automatic adjustments
			// [TODO-NExcel_Next]
//			format.setTimeZone(gmtZone);
			
			// Convert this to the number of days since 01 Jan 1970
			int offsetDays = nf?utcOffsetDays1904:utcOffsetDays;
			double utcDays = numValue - offsetDays;

			// Convert this into utc by multiplying by the number of milliseconds
			// in a day
//			long utcValue = (long) System.Math.Round(utcDays * msInADay);
			// convert it to 100 nanoseconds
			long utcValue = (long) System.Math.Round(utcDays * msInADay * 10000);

			// add the reference date (1/1/1970)
			DateTime refdate = new DateTime(1970, 1, 1);
			utcValue += refdate.Ticks;

			_Value = new DateTime(utcValue);
		}

		static DateRecord()
		{
			// [TODO-NExcel_Next]
//			gmtZone = TimeZone.CurrentTimeZone;
		}
	}
}
