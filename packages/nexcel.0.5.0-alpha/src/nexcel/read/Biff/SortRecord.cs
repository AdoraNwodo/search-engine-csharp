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
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	/// <summary> A storage area for the last Sort dialog box area</summary>
	public class SortRecord:RecordData
	{
		/// <summary> Accessor for the 1st Sort Column Name
		/// 
		/// </summary>
		/// <returns> the 1st Sort Column Name
		/// </returns>
		virtual public string SortCol1Name
		{
			get
			{
				return col1Name;
			}
			
		}
		/// <summary> Accessor for the 2nd Sort Column Name
		/// 
		/// </summary>
		/// <returns> the 2nd Sort Column Name
		/// </returns>
		virtual public string SortCol2Name
		{
			get
			{
				return col2Name;
			}
			
		}
		/// <summary> Accessor for the 3rd Sort Column Name
		/// 
		/// </summary>
		/// <returns> the 3rd Sort Column Name
		/// </returns>
		virtual public string SortCol3Name
		{
			get
			{
				return col3Name;
			}
			
		}
		/// <summary> Accessor for the Sort by Columns flag
		/// 
		/// </summary>
		/// <returns> the Sort by Columns flag
		/// </returns>
		virtual public bool SortColumns
		{
			get
			{
				return sortColumns;
			}
			
		}
		/// <summary> Accessor for the Sort Column 1 Descending flag
		/// 
		/// </summary>
		/// <returns> the Sort Column 1 Descending flag
		/// </returns>
		virtual public bool SortKey1Desc
		{
			get
			{
				return sortKey1Desc;
			}
			
		}
		/// <summary> Accessor for the Sort Column 2 Descending flag
		/// 
		/// </summary>
		/// <returns> the Sort Column 2 Descending flag
		/// </returns>
		virtual public bool SortKey2Desc
		{
			get
			{
				return sortKey2Desc;
			}
			
		}
		/// <summary> Accessor for the Sort Column 3 Descending flag
		/// 
		/// </summary>
		/// <returns> the Sort Column 3 Descending flag
		/// </returns>
		virtual public bool SortKey3Desc
		{
			get
			{
				return sortKey3Desc;
			}
			
		}
		/// <summary> Accessor for the Sort Case Sensitivity flag
		/// 
		/// </summary>
		/// <returns> the Sort Case Secsitivity flag
		/// </returns>
		virtual public bool SortCaseSensitive
		{
			get
			{
				return sortCaseSensitive;
			}
			
		}
		private int col1Size;
		private int col2Size;
		private int col3Size;
		private string col1Name;
		private string col2Name;
		private string col3Name;
		private sbyte optionFlags;
		private bool sortColumns = false;
		private bool sortKey1Desc = false;
		private bool sortKey2Desc = false;
		private bool sortKey3Desc = false;
		private bool sortCaseSensitive = false;
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="r">the raw data
		/// </param>
		public SortRecord(Record r):base(NExcel.Biff.Type.SORT)
		{
			
			sbyte[] data = r.Data;
			
			optionFlags = data[0];
			
			sortColumns = ((optionFlags & 0x01) != 0);
			sortKey1Desc = ((optionFlags & 0x02) != 0);
			sortKey2Desc = ((optionFlags & 0x04) != 0);
			sortKey3Desc = ((optionFlags & 0x08) != 0);
			sortCaseSensitive = ((optionFlags & 0x10) != 0);
			
			// data[1] contains sort list index - not implemented...
			
			col1Size = data[2];
			col2Size = data[3];
			col3Size = data[4];
			int curPos = 5;
			if (data[curPos++] == 0x00)
			{
				char[] tmpChar;
				tmpChar = new char[data.Length];
				data.CopyTo(tmpChar, 0);
				col1Name = new string(tmpChar, curPos, col1Size);
				curPos += col1Size;
			}
			else
			{
				col1Name = StringHelper.getUnicodeString(data, col1Size, curPos);
				curPos += col1Size * 2;
			}
			
			if (col2Size > 0)
			{
				if (data[curPos++] == 0x00)
				{
					char[] tmpChar2;
					tmpChar2 = new char[data.Length];
					data.CopyTo(tmpChar2, 0);
					col2Name = new string(tmpChar2, curPos, col2Size);
					curPos += col2Size;
				}
				else
				{
					col2Name = StringHelper.getUnicodeString(data, col2Size, curPos);
					curPos += col2Size * 2;
				}
			}
			else
			{
				col2Name = "";
			}
			if (col3Size > 0)
			{
				if (data[curPos++] == 0x00)
				{
					char[] tmpChar3;
					tmpChar3 = new char[data.Length];
					data.CopyTo(tmpChar3, 0);
					col3Name = new string(tmpChar3, curPos, col3Size);
					curPos += col3Size;
				}
				else
				{
					col3Name = StringHelper.getUnicodeString(data, col3Size, curPos);
					curPos += col3Size * 2;
				}
			}
			else
			{
				col3Name = "";
			}
		}
	}
}
