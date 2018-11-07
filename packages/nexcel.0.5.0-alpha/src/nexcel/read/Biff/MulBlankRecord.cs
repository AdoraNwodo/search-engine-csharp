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
	
	/// <summary> Contains an array of Blank, formatted cells</summary>
	class MulBlankRecord:RecordData
	{
		/// <summary> Accessor for the row
		/// 
		/// </summary>
		/// <returns> the row of containing these blank numbers
		/// </returns>
		virtual public int Row
		{
			get
			{
				return row;
			}
			
		}
		/// <summary> The first column containing the blank numbers
		/// 
		/// </summary>
		/// <returns> the first column
		/// </returns>
		virtual public int FirstColumn
		{
			get
			{
				return colFirst;
			}
			
		}
		/// <summary> Accessor for the number of blank values
		/// 
		/// </summary>
		/// <returns> the number of blank values
		/// </returns>
		virtual public int NumberOfColumns
		{
			get
			{
				return numblanks;
			}
			
		}
		/// <summary> The row  containing these numbers</summary>
		private int row;
		/// <summary> The first column these rk number occur on</summary>
		private int colFirst;
		/// <summary> The last column these blank numbers occur on</summary>
		private int colLast;
		/// <summary> The number of blank numbers contained in this record</summary>
		private int numblanks;
		/// <summary> The array of xf indices</summary>
		private int[] xfIndices;
		
		/// <summary> Constructs the blank records from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public MulBlankRecord(Record t):base(t)
		{
			sbyte[] data = getRecord().Data;
			int length = getRecord().Length;
			row = IntegerHelper.getInt(data[0], data[1]);
			colFirst = IntegerHelper.getInt(data[2], data[3]);
			colLast = IntegerHelper.getInt(data[length - 2], data[length - 1]);
			numblanks = colLast - colFirst + 1;
			xfIndices = new int[numblanks];
			
			readBlanks(data);
		}
		
		/// <summary> Reads the blanks from the raw data
		/// 
		/// </summary>
		/// <param name="data">the raw data
		/// </param>
		private void  readBlanks(sbyte[] data)
		{
			int pos = 4;
			//    int blank;
			for (int i = 0; i < numblanks; i++)
			{
				xfIndices[i] = IntegerHelper.getInt(data[pos], data[pos + 1]);
				pos += 2;
			}
		}
		
		/// <summary> Return a specific formatting index</summary>
		/// <param name="index">the cell index in the group
		/// </param>
		/// <returns> the formatting index
		/// </returns>
		public virtual int getXFIndex(int index)
		{
			return xfIndices[index];
		}
	}
}
