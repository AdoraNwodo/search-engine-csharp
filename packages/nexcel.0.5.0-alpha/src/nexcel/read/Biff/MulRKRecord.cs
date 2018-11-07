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
	
	/// <summary> 
	/// Contains an array of RK numbers
	/// </summary>
	class MulRKRecord:RecordData
	{
		/// <summary> Accessor for the row
		/// 
		/// </summary>
		/// <returns> the row of containing these rk numbers
		/// </returns>
		virtual public int Row
		{
			get
			{
				return row;
			}
			
		}
		/// <summary> The first column containing the rk numbers
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
		/// <summary> Accessor for the number of rk values
		/// 
		/// </summary>
		/// <returns> the number of rk values
		/// </returns>
		virtual public int NumberOfColumns
		{
			get
			{
				return numrks;
			}
			
		}
		/// <summary> The row  containing these numbers</summary>
		private int row;
		/// <summary> The first column these rk number occur on</summary>
		private int colFirst;
		/// <summary> The last column these rk numbers occur on</summary>
		private int colLast;
		/// <summary> The number of rk numbers contained in this record</summary>
		private int numrks;
		/// <summary> The array of rk numbers</summary>
		private int[] rknumbers;
		/// <summary> The array of xf indices</summary>
		private int[] xfIndices;
		
		/// <summary> Constructs the rk numbers from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public MulRKRecord(Record t):base(t)
		{
			sbyte[] data = getRecord().Data;
			int length = getRecord().Length;
			row = IntegerHelper.getInt(data[0], data[1]);
			colFirst = IntegerHelper.getInt(data[2], data[3]);
			colLast = IntegerHelper.getInt(data[length - 2], data[length - 1]);
			numrks = colLast - colFirst + 1;
			rknumbers = new int[numrks];
			xfIndices = new int[numrks];
			
			readRks(data);
		}
		
		/// <summary> Reads the rks from the raw data
		/// 
		/// </summary>
		/// <param name="data">the raw data
		/// </param>
		private void  readRks(sbyte[] data)
		{
			int pos = 4;
			int rk;
			for (int i = 0; i < numrks; i++)
			{
				xfIndices[i] = IntegerHelper.getInt(data[pos], data[pos + 1]);
				rk = IntegerHelper.getInt(data[pos + 2], data[pos + 3], data[pos + 4], data[pos + 5]);
				rknumbers[i] = rk;
				pos += 6;
			}
		}
		
		/// <summary> Returns a specific rk number
		/// 
		/// </summary>
		/// <param name="index">the rk number to return
		/// </param>
		/// <returns> the rk number in bits
		/// </returns>
		public virtual int getRKNumber(int index)
		{
			return rknumbers[index];
		}
		
		/// <summary> Return a specific formatting index
		/// 
		/// </summary>
		/// <param name="index">the index of the cell in this group
		/// </param>
		/// <returns> the xf index
		/// </returns>
		public virtual int getXFIndex(int index)
		{
			return xfIndices[index];
		}
	}
}
