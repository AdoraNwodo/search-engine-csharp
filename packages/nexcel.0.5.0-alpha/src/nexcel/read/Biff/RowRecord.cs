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
	
	/// <summary> A row  record</summary>
	public class RowRecord:RecordData
	{
		/// <summary> Interrogates whether this row is of default height
		/// 
		/// </summary>
		/// <returns> TRUE if this is set to the default height, FALSE otherwise
		/// </returns>
		virtual internal bool isDefaultHeight()
		{
				return rowHeight == defaultHeightIndicator;
		}
		/// <summary> Gets the row number
		/// 
		/// </summary>
		/// <returns> the number of this row
		/// </returns>
		virtual public int RowNumber
		{
			get
			{
				return rowNumber;
			}
			
		}
		/// <summary> Gets the height of the row
		/// 
		/// </summary>
		/// <returns> the row height
		/// </returns>
		virtual public int RowHeight
		{
			get
			{
				return rowHeight;
			}
			
		}
		/// <summary> Queries whether the row is collapsed
		/// 
		/// </summary>
		/// <returns> the collapsed indicator
		/// </returns>
		virtual public bool isCollapsed()
		{
				return collapsed;
		}
		/// <summary> Queries whether the row has been set to zero height
		/// 
		/// </summary>
		/// <returns> the zero height indicator
		/// </returns>
		virtual public bool isZeroHeight()
		{
				return zeroHeight;
		}

		/// <summary> The number of this row</summary>
		private int rowNumber;
		/// <summary> The height of this row</summary>
		private int rowHeight;
		/// <summary> Flag to indicate whether this row is collapsed or not</summary>
		private bool collapsed;
		/// <summary> Indicates whether this row has zero height (ie. whether it is hidden)</summary>
		private bool zeroHeight;
		
		/// <summary> Indicates that the row is default height</summary>
		private const int defaultHeightIndicator = 0xff;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		internal RowRecord(Record t):base(t)
		{
			
			sbyte[] data = getRecord().Data;
			rowNumber = IntegerHelper.getInt(data[0], data[1]);
			rowHeight = IntegerHelper.getInt(data[6], data[7]);
			
			sbyte opts = data[12];
			
			collapsed = (opts & 0x20) != 0;
		}
	}
}
