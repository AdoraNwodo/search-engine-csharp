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
using NExcel;
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	/// <summary> A record containing the references to the various sheets (internal and
	/// external) referenced by formulas in this workbook
	/// </summary>
	public class SupbookRecord:RecordData
	{
		/// <summary> Gets the number of sheets.  This will only be non-zero for internal
		/// and external supbooks
		/// 
		/// </summary>
		/// <returns> the number of sheets
		/// </returns>
		virtual public int NumberOfSheets
		{
			get
			{
				return numSheets;
			}
			
		}
		/// <summary> Gets the name of the external file
		/// 
		/// </summary>
		/// <returns> the name of the external file
		/// </returns>
		virtual public string FileName
		{
			get
			{
				return fileName;
			}
			
		}
		/// <summary> Gets the data - used when copying a spreadsheet
		/// 
		/// </summary>
		/// <returns> the raw external sheet data
		/// </returns>
		virtual public sbyte[] Data
		{
			get
			{
				return getRecord().Data;
			}
			
		}
		/// <summary> The type of this supbook record</summary>
		private SupbookRecord.SupbookType type;
		
		/// <summary> The number of sheets - internal & external supbooks only</summary>
		private int numSheets;
		
		/// <summary> The name of the external file</summary>
		private string fileName;
		
		/// <summary> The names of the external sheets</summary>
		private string[] sheetNames;
		
		/// <summary> The type of supbook this refers to</summary>
		public class SupbookType
		{
		}
		
		
		public static readonly SupbookType INTERNAL = new SupbookType();
		public static readonly SupbookType EXTERNAL = new SupbookType();
		public static readonly SupbookType ADDIN = new SupbookType();
		public static readonly SupbookType LINK = new SupbookType();
		public static readonly SupbookType UNKNOWN = new SupbookType();
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		internal SupbookRecord(Record t, WorkbookSettings ws):base(t)
		{
			sbyte[] data = getRecord().Data;
			
			// First deduce the type
			if (data.Length == 4)
			{
				if (data[2] == 0x01 && data[3] == 0x04)
				{
					type = INTERNAL;
				}
				else if (data[2] == 0x01 && data[3] == 0x3a)
				{
					type = ADDIN;
				}
				else
				{
					type = UNKNOWN;
				}
			}
			else if (data[0] == 0 && data[1] == 0)
			{
				type = LINK;
			}
			else
			{
				type = EXTERNAL;
			}
			
			if (type == INTERNAL)
			{
				numSheets = IntegerHelper.getInt(data[0], data[1]);
			}
			
			if (type == EXTERNAL)
			{
				readExternal(data, ws);
			}
		}
		
		/// <summary> Reads the external data records
		/// 
		/// </summary>
		/// <param name="data">the data
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		private void  readExternal(sbyte[] data, WorkbookSettings ws)
		{
			numSheets = IntegerHelper.getInt(data[0], data[1]);
			
			// subtract file name encoding from the .Length
			int ln = IntegerHelper.getInt(data[2], data[3]) - 1;
			int pos = 0;
			
			if (data[4] == 0)
			{
				// non-unicode string
				int encoding = data[5];
				pos = 6;
				if (encoding == 0)
				{
					fileName = StringHelper.getString(data, ln, pos, ws);
					pos += ln;
				}
				else
				{
					fileName = getEncodedFilename(data, ln, pos);
					pos += ln;
				}
			}
			else
			{
				// unicode string
				int encoding = IntegerHelper.getInt(data[5], data[6]);
				pos = 7;
				if (encoding == 0)
				{
					fileName = StringHelper.getUnicodeString(data, ln, pos);
					pos += ln * 2;
				}
				else
				{
					fileName = getUnicodeEncodedFilename(data, ln, pos);
					pos += ln * 2;
				}
			}
			
			sheetNames = new string[numSheets];
			
			for (int i = 0; i < sheetNames.Length; i++)
			{
				ln = IntegerHelper.getInt(data[pos], data[pos + 1]);
				
				if (data[pos + 2] == 0x0)
				{
					sheetNames[i] = StringHelper.getString(data, ln, pos + 3, ws);
					pos += ln + 3;
				}
				else if (data[pos + 2] == 0x1)
				{
					sheetNames[i] = StringHelper.getUnicodeString(data, ln, pos + 3);
					pos += ln * 2 + 3;
				}
			}
		}
		
		/// <summary> Gets the type of this supbook record
		/// 
		/// </summary>
		/// <returns> the type of this supbook
		/// </returns>
		public virtual SupbookType Type
		{
			get
			{
				return type;
			}
		}
		
		/// <summary> Gets the name of the external sheet
		/// 
		/// </summary>
		/// <param name="i">the index of the external sheet
		/// </param>
		/// <returns> the name of the sheet
		/// </returns>
		public virtual string getSheetName(int i)
		{
			return sheetNames[i];
		}
		
		/// <summary> Gets the encoded string from the data array
		/// 
		/// </summary>
		/// <param name="data">the data
		/// </param>
		/// <param name="ln">.Length of the string
		/// </param>
		/// <param name="pos">the position in the array
		/// </param>
		/// <returns> the string
		/// </returns>
		private string getEncodedFilename(sbyte[] data, int ln, int pos)
		{
			System.Text.StringBuilder buf = new System.Text.StringBuilder();
			int endpos = pos + ln;
			while (pos < endpos)
			{
				char c = (char) data[pos];
				
				if (c == '\u0001')
				{
					// next character is a volume letter
					pos++;
					c = (char) data[pos];
					buf.Append(c);
					buf.Append(":\\\\");
				}
				else if (c == '\u0002')
				{
					// file is on the same volume
					buf.Append('\\');
				}
				else if (c == '\u0003')
				{
					// down directory
					buf.Append('\\');
				}
				else if (c == '\u0004')
				{
					// up directory
					buf.Append("..\\");
				}
				else
				{
					// just add on the character
					buf.Append(c);
				}
				
				pos++;
			}
			
			return buf.ToString();
		}
		
		/// <summary> Gets the encoded string from the data array
		/// 
		/// </summary>
		/// <param name="data">the data
		/// </param>
		/// <param name="ln">.Length of the string
		/// </param>
		/// <param name="pos">the position in the array
		/// </param>
		/// <returns> the string
		/// </returns>
		private string getUnicodeEncodedFilename(sbyte[] data, int ln, int pos)
		{
			System.Text.StringBuilder buf = new System.Text.StringBuilder();
			int endpos = pos + ln * 2;
			while (pos < endpos)
			{
				char c = (char) IntegerHelper.getInt(data[pos], data[pos + 1]);
				
				if (c == '\u0001')
				{
					// next character is a volume letter
					pos += 2;
					c = (char) IntegerHelper.getInt(data[pos], data[pos + 1]);
					buf.Append(c);
					buf.Append(":\\\\");
				}
				else if (c == '\u0002')
				{
					// file is on the same volume
					buf.Append('\\');
				}
				else if (c == '\u0003')
				{
					// down directory
					buf.Append('\\');
				}
				else if (c == '\u0004')
				{
					// up directory
					buf.Append("..\\");
				}
				else
				{
					// just add on the character
					buf.Append(c);
				}
				
				pos += 2;
			}
			
			return buf.ToString();
		}
	}
}
