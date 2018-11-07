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
using System.Text;
using NExcelUtils;
using common;
using NExcel;
using NExcel.Read.Biff;
using NExcel.Format;
namespace NExcel.Biff
{
	
	/// <summary> A non-built format record</summary>
	public class FormatRecord:WritableRecordData, DisplayFormat, NExcel.Format.Format
	{
		/// <summary> Gets the format index of this record
		/// 
		/// </summary>
		/// <returns> the format index of this record
		/// </returns>
		virtual public int FormatIndex
		{
			get
			{
				return indexCode;
			}
			
		}
		/// <summary> Accessor to see whether this object is initialized or not.
		/// 
		/// </summary>
		/// <returns> TRUE if this font record has been initialized, FALSE otherwise
		/// </returns>
		virtual public bool isInitialized()
		{
				return initialized;
		}

		/// <summary> Sees if this format is a date format
		/// 
		/// </summary>
		/// <returns> TRUE if this format is a date
		/// </returns>
		virtual public bool Date
		{
			get
			{
				return date;
			}
			
		}
		/// <summary> Sees if this format is a number format
		/// 
		/// </summary>
		/// <returns> TRUE if this format is a number
		/// </returns>
		virtual public bool Number
		{
			get
			{
				return number;
			}
			
		}
		/// <summary> Gets the java equivalent number format for the formatString
		/// 
		/// </summary>
		/// <returns> The java equivalent of the number format for this object
		/// </returns>
		virtual public NumberFormatInfo NumberFormat
		{
			get
			{
				if (format != null && format is NumberFormatInfo)
				{
					return (NumberFormatInfo) format;
				}
				
				try
				{
					string fs = formatString;
					
					// Replace the Excel formatting characters with java equivalents
					fs = replace(fs, "E+", "E");
					fs = replace(fs, "_)", "");
					fs = replace(fs, "_", "");
					fs = replace(fs, "[Red]", "");
					fs = replace(fs, "\\", "");
					
					format = new NumberFormatInfo(fs);
				}
				catch (System.ArgumentException e)
				{
					// Something went wrong with the date format - fail silently
					// and return a default value
					format = new NumberFormatInfo("#.###");
				}
				
				return (NumberFormatInfo) format;
			}
			
		}
		/// <summary> Gets the java equivalent date format for the formatString
		/// 
		/// </summary>
		/// <returns> The lang-specific equivalent of the date format for this object
		/// </returns>
		virtual public DateTimeFormatInfo DateFormat
		{
			get
			{
				if (format != null && format is DateTimeFormatInfo)
				{
					return (DateTimeFormatInfo) format;
				}
				
				string fmt = formatString;

				// [TODO-NExcel_Next] - bad solution, do better way in future
				if (fmt.EndsWith(";@")) fmt = fmt.Substring(0, fmt.Length - 2);
				if (fmt.StartsWith("[$-"))
				{
					int iend = fmt.IndexOf("]", 0);
					fmt = fmt.Substring(iend+1);
				}
				
				
				// Replace the AM/PM indicator with an a
				int pos = fmt.IndexOf("AM/PM");
				while (pos != - 1)
				{
					StringBuilder csb = new StringBuilder(fmt.Substring(0, pos));
//					csb.Append('a');
					csb.Append("tt");
					csb.Append(fmt.Substring(pos + 5));
					fmt = csb.ToString();
					pos = fmt.IndexOf("AM/PM");
				}
				
				// Replace ss.0 with ss.SSS (necessary to always specify milliseconds
				// because of NT)
				pos = fmt.IndexOf("ss.0");
				while (pos != - 1)
				{
					System.Text.StringBuilder csb = new System.Text.StringBuilder(fmt.Substring(0, (pos) - (0)));
//					csb.Append("ss.SSS");
					csb.Append("ss.fff");
					
					// Keep going until we run out of zeros
					pos += 4;
					while (pos < fmt.Length && fmt[pos] == '0')
					{
						pos++;
					}
					
					csb.Append(fmt.Substring(pos));
					fmt = csb.ToString();
					pos = fmt.IndexOf("ss.0");
				}
				
				
				// Filter out the backslashes
				StringBuilder sb = new StringBuilder();
				for (int i = 0; i < fmt.Length; i++)
				{
					if (fmt[i] != '\\')
					{
						sb.Append(fmt[i]);
					}
				}
				
				fmt = sb.ToString();
				
				// We need to convert the month indicator m, to upper case when we
				// are dealing with dates
				char[] formatBytes = fmt.ToCharArray();
				
				for (int i = 0; i < formatBytes.Length; i++)
				{
					if (formatBytes[i] == 'm')
					{
						// Firstly, see if the preceding character is also an m.  If so,
						// copy that
						if (i > 0 && (formatBytes[i - 1] == 'm' || formatBytes[i - 1] == 'M'))
						{
							formatBytes[i] = formatBytes[i - 1];
						}
						else
						{
							// There is no easy way out.  We have to deduce whether this an
							// minute or a month?  See which is closest out of the
							// letters H d s or y
							// First, h
							int minuteDist = System.Int32.MaxValue;
							for (int j = i - 1; j >= 0; j--)
							{
								if (formatBytes[j] == 'h')
								{
									minuteDist = i - j;
									break;
								}
							}
							
							for (int j = i + 1; j < formatBytes.Length; j++)
							{
								if (formatBytes[j] == 'h')
								{
									minuteDist = System.Math.Min(minuteDist, j - i);
									break;
								}
							}
							
							for (int j = i - 1; j >= 0; j--)
							{
								if (formatBytes[j] == 'H')
								{
									minuteDist = i - j;
									break;
								}
							}
							
							for (int j = i + 1; j < formatBytes.Length; j++)
							{
								if (formatBytes[j] == 'H')
								{
									minuteDist = System.Math.Min(minuteDist, j - i);
									break;
								}
							}
							
							// Now repeat for s
							for (int j = i - 1; j >= 0; j--)
							{
								if (formatBytes[j] == 's')
								{
									minuteDist = System.Math.Min(minuteDist, i - j);
									break;
								}
							}
							for (int j = i + 1; j < formatBytes.Length; j++)
							{
								if (formatBytes[j] == 's')
								{
									minuteDist = System.Math.Min(minuteDist, j - i);
									break;
								}
							}
							// We now have the distance of the closest character which could
							// indicate the the m refers to a minute
							// Repeat for d and y
							int monthDist = System.Int32.MaxValue;
							for (int j = i - 1; j >= 0; j--)
							{
								if (formatBytes[j] == 'd')
								{
									monthDist = i - j;
									break;
								}
							}
							
							for (int j = i + 1; j < formatBytes.Length; j++)
							{
								if (formatBytes[j] == 'd')
								{
									monthDist = System.Math.Min(monthDist, j - i);
									break;
								}
							}
							// Now repeat for y
							for (int j = i - 1; j >= 0; j--)
							{
								if (formatBytes[j] == 'y')
								{
									monthDist = System.Math.Min(monthDist, i - j);
									break;
								}
							}
							for (int j = i + 1; j < formatBytes.Length; j++)
							{
								if (formatBytes[j] == 'y')
								{
									monthDist = System.Math.Min(monthDist, j - i);
									break;
								}
							}
							
							if (monthDist < minuteDist)
							{
								// The month indicator is closer, so convert to a capital M
								formatBytes[i] = System.Char.ToUpper(formatBytes[i]);
							}
							else if ((monthDist == minuteDist) && (monthDist != System.Int32.MaxValue))
							{
								// They are equidistant.  As a tie-breaker, take the formatting
								// character which precedes the m
								char ind = formatBytes[i - monthDist];
								if (ind == 'y' || ind == 'd')
								{
									// The preceding item indicates a month measure, so convert
									formatBytes[i] = System.Char.ToUpper(formatBytes[i]);
								}
							}
						}
					}
				}
				
				try
				{
					this.format = new DateTimeFormatInfo(new string(formatBytes));
				}
				catch (System.ArgumentException /*e*/)
				{
					// There was a spurious character - fail silently
					this.format = new DateTimeFormatInfo("dd MM yyyy hh:mm:ss");
				}
				return (DateTimeFormatInfo) this.format;
			}
			
		}
		/// <summary> Gets the index code, for use as a hash value
		/// 
		/// </summary>
		/// <returns> the ifmt code for this cell
		/// </returns>
		virtual public int IndexCode
		{
			get
			{
				return indexCode;
			}
			
		}
		/// <summary> Indicates whether this formula is a built in
		/// 
		/// </summary>
		/// <returns> FALSE
		/// </returns>
		virtual public bool isBuiltIn()
		{
				return false;
		}


		/// <summary> The logger</summary>
		public static Logger logger;
		
		/// <summary> Initialized flag</summary>
		private bool initialized;
		
		/// <summary> The raw data</summary>
		private sbyte[] data;
		
		/// <summary> The index code</summary>
		private int indexCode;
		
		/// <summary> The formatting string</summary>
		private string formatString;
		
		/// <summary> Indicates whether this is a date formatting record</summary>
		private bool date;
		
		/// <summary> Indicates whether this a number formatting record</summary>
		private bool number;
		
		/// <summary> The format object</summary>
		private IFormatProvider format;
		
		/// <summary> The date strings to look for</summary>
		private static string[] dateStrings = new string[]{"dd", "mm", "yy", "hh", "ss", "m/", "/d"};
		
		// Type to distinguish between biff7 and biff8
		public class BiffType
		{
		}
		
		public static readonly BiffType biff8 = new BiffType();
		public static readonly BiffType biff7 = new BiffType();
		
		/// <summary> Constructor invoked when copying sheets
		/// 
		/// </summary>
		/// <param name="fmt">the format string
		/// </param>
		/// <param name="refno">the index code
		/// </param>
		internal FormatRecord(string fmt, int refno):base(NExcel.Biff.Type.FORMAT)
		{
			formatString = fmt;
			indexCode = refno;
			initialized = true;
		}
		
		/// <summary> Constructor used by writable formats</summary>
		protected internal FormatRecord():base(NExcel.Biff.Type.FORMAT)
		{
			initialized = false;
		}
		
		/// <summary> Copy constructor - can be invoked by public access
		/// 
		/// </summary>
		/// <param name="fr">the format to copy
		/// </param>
		protected internal FormatRecord(FormatRecord fr):base(NExcel.Biff.Type.FORMAT)
		{
			initialized = false;
			
			formatString = fr.formatString;
			date = fr.date;
			number = fr.number;
			//    format = (java.text.Format) fr.format.clone();
		}
		
		/// <summary> Constructs this object from the raw data.  Used when reading in a
		/// format record
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <param name="biffType">biff type dummy overload
		/// </param>
		public FormatRecord(Record t, WorkbookSettings ws, BiffType biffType):base(t)
		{
			
			sbyte[] data = getRecord().Data;
			indexCode = IntegerHelper.getInt(data[0], data[1]);
			initialized = true;
			
			if (biffType == biff8)
			{
				int numchars = IntegerHelper.getInt(data[2], data[3]);
				if (data[4] == 0)
				{
					formatString = StringHelper.getString(data, numchars, 5, ws);
				}
				else
				{
					formatString = StringHelper.getUnicodeString(data, numchars, 5);
				}
			}
			else
			{
				int numchars = data[2];
				sbyte[] chars = new sbyte[numchars];
				Array.Copy(data, 3, chars, 0, chars.Length);
				formatString = new string(NExcelUtils.Byte.ToCharArray(NExcelUtils.Byte.ToByteArray(chars)));
			}
			
			date = false;
			number = false;
			
			// First see if this is a date format
			for (int i = 0; i < dateStrings.Length; i++)
			{
				string dateString = dateStrings[i];
				if (formatString.IndexOf(dateString) != - 1 || formatString.IndexOf(dateString.ToUpper()) != - 1)
				{
					date = true;
					break;
				}
			}
			
			// See if this is number format - look for the # or 0 characters
			if (!date)
			{
				if (formatString.IndexOf((System.Char) '#') != - 1 || formatString.IndexOf((System.Char) '0') != - 1)
				{
					number = true;
				}
			}
		}
		
		/// <summary> Used to get the data when writing out the format record
		/// 
		/// </summary>
		/// <returns> the raw data
		/// </returns>
		public override sbyte[] getData()
		{
			data = new sbyte[formatString.Length * 2 + 3 + 2];
			
			IntegerHelper.getTwoBytes(indexCode, data, 0);
			IntegerHelper.getTwoBytes(formatString.Length, data, 2);
			data[4] = (sbyte) 1; // unicode indicator
			StringHelper.getUnicodeBytes(formatString, data, 5);
			
			return data;
		}
		
		/// <summary> Sets the index of this record.  Called from the FormattingRecords
		/// object
		/// 
		/// </summary>
		/// <param name="pos">the position of this font in the workbooks font list
		/// </param>
		
		public virtual void  initialize(int pos)
		{
			indexCode = pos;
			initialized = true;
		}
		
		/// <summary> Replaces all instances of search with replace in the input.  Used for
		/// replacing microsoft number formatting characters with java equivalents
		/// 
		/// </summary>
		/// <param name="input">the format string
		/// </param>
		/// <param name="search">the Excel character to be replaced
		/// </param>
		/// <param name="replace">the java equivalent
		/// </param>
		/// <returns> the input string with the specified substring replaced
		/// </returns>
		protected internal string replace(string input, string search, string replace)
		{
			string fmtstr = input;
			int pos = fmtstr.IndexOf(search);
			while (pos != - 1)
			{
				System.Text.StringBuilder tmp = new System.Text.StringBuilder(fmtstr.Substring(0, (pos) - (0)));
				tmp.Append(replace);
				tmp.Append(fmtstr.Substring(pos + search.Length));
				fmtstr = tmp.ToString();
				pos = fmtstr.IndexOf(search);
			}
			return fmtstr;
		}
		
		/// <summary> Called by the immediate subclass to set the string
		/// once the Java-Excel replacements have been done
		/// 
		/// </summary>
		/// <param name="s">the format string
		/// </param>
		protected internal void  setFormatString(string s)
		{
			formatString = s;
		}
		
		/// <summary> Gets the formatting string.
		/// 
		/// </summary>
		/// <returns> the excel format string
		/// </returns>
		public virtual string FormatString
		{
			get
			{
				return formatString;
			}
		}
		
		/// <summary> Standard hash code method</summary>
		/// <returns> the hash code value for this object
		/// </returns>
		public override int GetHashCode()
		{
			return formatString.GetHashCode();
		}
		
		/// <summary> Standard equals method.  This compares the contents of two
		/// format records, and not their indexCodes, which are ignored
		/// 
		/// </summary>
		/// <param name="o">the object to compare
		/// </param>
		/// <returns> TRUE if the two objects are equal, FALSE otherwise
		/// </returns>
		public  override bool Equals(System.Object o)
		{
			if (o == this)
			{
				return true;
			}
			
			if (!(o is FormatRecord))
			{
				return false;
			}
			
			FormatRecord fr = (FormatRecord) o;
			
			// Not interested in uninitialized comparisons
			if (!initialized || !fr.initialized)
			{
				return false;
			}
			
			// Must be either a number or a date format
			if (date != fr.date || number != fr.number)
			{
				return false;
			}
			
			return formatString.Equals(fr.formatString);
		}
		static FormatRecord()
		{
			logger = Logger.getLogger(typeof(FormatRecord));
		}
	}
}
