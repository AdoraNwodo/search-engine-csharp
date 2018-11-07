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
	
	/// <summary> Contains the cell dimensions of this worksheet</summary>
	class HorizontalPageBreaksRecord:RecordData
	{
		/// <summary> Gets the row breaks
		/// 
		/// </summary>
		/// <returns> the row breaks on the current sheet
		/// </returns>
		virtual public int[] RowBreaks
		{
			get
			{
				return rowBreaks;
			}
			
		}
		/// <summary> The row page breaks</summary>
		private int[] rowBreaks;
		
		/// <summary> Dummy indicators for overloading the constructor</summary>
		public class Biff7
		{
		}
		
		public static Biff7 biff7;
		
		/// <summary> Constructs the dimensions from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public HorizontalPageBreaksRecord(Record t):base(t)
		{
			
			sbyte[] data = t.Data;
			
			int numbreaks = IntegerHelper.getInt(data[0], data[1]);
			int pos = 2;
			rowBreaks = new int[numbreaks];
			
			for (int i = 0; i < numbreaks; i++)
			{
				rowBreaks[i] = IntegerHelper.getInt(data[pos], data[pos + 1]);
				pos += 6;
			}
		}
		
		/// <summary> Constructs the dimensions from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="biff7">an indicator to initialise this record for biff 7 format
		/// </param>
		public HorizontalPageBreaksRecord(Record t, Biff7 biff7):base(t)
		{
			
			sbyte[] data = t.Data;
			int numbreaks = IntegerHelper.getInt(data[0], data[1]);
			int pos = 2;
			rowBreaks = new int[numbreaks];
			
			for (int i = 0; i < numbreaks; i++)
			{
				pos += 2;
				rowBreaks[i] = IntegerHelper.getInt(data[pos], data[pos + 1]);
			}
		}
		static HorizontalPageBreaksRecord()
		{
			biff7 = new Biff7();
		}
	}
}
