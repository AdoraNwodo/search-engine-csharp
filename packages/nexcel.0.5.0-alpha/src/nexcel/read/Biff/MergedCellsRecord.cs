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
	
	/// <summary> A merged cells record for a given sheet</summary>
	public class MergedCellsRecord:RecordData
	{
		/// <summary> Gets the ranges which have been merged in this sheet
		/// 
		/// </summary>
		/// <returns> the ranges of cells which have been merged
		/// </returns>
		virtual public Range[] Ranges
		{
			get
			{
				return ranges;
			}
			
		}
		/// <summary> The ranges of the cells merged on this sheet</summary>
		private Range[] ranges;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="s">the sheet
		/// </param>
		internal MergedCellsRecord(Record t, Sheet s):base(t)
		{
			
			sbyte[] data = getRecord().Data;
			
			int numRanges = IntegerHelper.getInt(data[0], data[1]);
			
			ranges = new Range[numRanges];
			
			int pos = 2;
			int firstRow = 0;
			int lastRow = 0;
			int firstCol = 0;
			int lastCol = 0;
			
			for (int i = 0; i < numRanges; i++)
			{
				firstRow = IntegerHelper.getInt(data[pos], data[pos + 1]);
				lastRow = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
				firstCol = IntegerHelper.getInt(data[pos + 4], data[pos + 5]);
				lastCol = IntegerHelper.getInt(data[pos + 6], data[pos + 7]);
				
				ranges[i] = new SheetRangeImpl(s, firstCol, firstRow, lastCol, lastRow);
				
				pos += 8;
			}
		}
	}
}
