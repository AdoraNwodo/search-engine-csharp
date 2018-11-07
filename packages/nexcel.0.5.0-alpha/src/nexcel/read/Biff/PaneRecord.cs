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
	class PaneRecord:RecordData
	{
		/// <summary> Accessor for the number of rows in the top left pane
		/// 
		/// </summary>
		/// <returns> the number of rows visible in the top left pane
		/// </returns>
		virtual public int RowsVisible
		{
			get
			{
				return rowsVisible;
			}
			
		}
		/// <summary> Accessor for the numbe rof columns visible in the top left pane
		/// 
		/// </summary>
		/// <returns> the number of columns visible in the top left pane
		/// </returns>
		virtual public int ColumnsVisible
		{
			get
			{
				return columnsVisible;
			}
			
		}
		/// <summary> The number of rows visible in the top left pane</summary>
		private int rowsVisible;
		/// <summary> The number of columns visible in the top left pane</summary>
		private int columnsVisible;
		
		/// <summary> Constructs the dimensions from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public PaneRecord(Record t):base(t)
		{
			sbyte[] data = t.Data;
			
			columnsVisible = IntegerHelper.getInt(data[0], data[1]);
			rowsVisible = IntegerHelper.getInt(data[2], data[3]);
		}
	}
}
