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
using DateTime = System.DateTime;
using DateTimeFormatInfo = NExcelUtils.DateTimeFormatInfo;
namespace NExcel
{
	
	/// <summary> A date cell</summary>
	public interface DateCell : Cell
		{
			/// <summary> Gets the date contained in this cell
			/// 
			/// </summary>
			/// <returns> the cell contents
			/// </returns>
			DateTime DateValue
			{
				get;
				
			}
			/// <summary> Indicates whether the date value contained in this cell refers to a date,
			/// or merely a time
			/// 
			/// </summary>
			/// <returns> TRUE if the value refers to a time
			/// </returns>
			bool Time
			{
				get;
				
			}
			/// <summary> Gets the DateFormat used to format the cell.  This will normally be
			/// the format specified in the excel spreadsheet, but in the event of any
			/// difficulty parsing this, it will revert to the default date/time format.
			/// 
			/// </summary>
			/// <returns> the DateFormat object used to format the date in the original
			/// excel cell
			/// </returns>
			DateTimeFormatInfo DateFormat
			{
				get;
				
			}
		}
}
