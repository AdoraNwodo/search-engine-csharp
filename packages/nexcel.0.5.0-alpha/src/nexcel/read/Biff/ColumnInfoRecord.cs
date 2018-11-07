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
	
	/// <summary> Contains the display info data which affects the entire columns</summary>
	public class ColumnInfoRecord:RecordData
	{
		/// <summary> Accessor for the start column of this range
		/// 
		/// </summary>
		/// <returns> the start column index
		/// </returns>
		virtual public int StartColumn
		{
			get
			{
				return startColumn;
			}
			
		}
		/// <summary> Accessor for the end column of this range
		/// 
		/// </summary>
		/// <returns> the end column index
		/// </returns>
		virtual public int EndColumn
		{
			get
			{
				return endColumn;
			}
			
		}
		/// <summary> Accessor for the column format index
		/// 
		/// </summary>
		/// <returns> the format index
		/// </returns>
		virtual public int XFIndex
		{
			get
			{
				return xfIndex;
			}
			
		}
		/// <summary> Accessor for the width of the column
		/// 
		/// </summary>
		/// <returns> the width
		/// </returns>
		virtual public int Width
		{
			get
			{
				return width;
			}
			
		}
		/// <summary> Accessor for the hidden flag. Used when copying sheets
		/// 
		/// </summary>
		/// <returns> TRUE if the columns are hidden, FALSE otherwise
		/// </returns>
		virtual public bool Hidden
		{
			get
			{
				return hidden;
			}
			
		}
		/// <summary> The raw data</summary>
		private sbyte[] data;
		
		/// <summary> The start for which to apply the format information</summary>
		private int startColumn;
		
		/// <summary> The end column for which to apply the format information</summary>
		private int endColumn;
		
		/// <summary> The index to the XF record, which applies to each cell in this column</summary>
		private int xfIndex;
		
		/// <summary> The width of the column in 1/256 of a character</summary>
		private int width;
		
		/// <summary> A hidden flag</summary>
		private bool hidden;
		
		/// <summary> Constructor which creates this object from the binary data
		/// 
		/// </summary>
		/// <param name="t">the record
		/// </param>
		internal ColumnInfoRecord(Record t):base(NExcel.Biff.Type.COLINFO)
		{
			
			data = t.Data;
			
			startColumn = IntegerHelper.getInt(data[0], data[1]);
			endColumn = IntegerHelper.getInt(data[2], data[3]);
			width = IntegerHelper.getInt(data[4], data[5]);
			xfIndex = IntegerHelper.getInt(data[6], data[7]);
			
			int options = IntegerHelper.getInt(data[8], data[9]);
			hidden = ((options & 0x1) != 0);
		}
	}
}
