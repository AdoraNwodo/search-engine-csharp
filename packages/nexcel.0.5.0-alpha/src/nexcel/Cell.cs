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
using NExcel.Format;
namespace NExcel
{
	
	/// <summary> Represents an individual Cell within a Sheet.  May be queried for its
	/// type and its content
	/// </summary>
	public interface Cell
		{
			/// <summary> Returns the row number of this cell
			/// 
			/// </summary>
			/// <returns> the row number of this cell
			/// </returns>
			int Row
			{
				get;
				
			}
			/// <summary> Returns the column number of this cell
			/// 
			/// </summary>
			/// <returns> the column number of this cell
			/// </returns>
			int Column
			{
				get;
				
			}
			/// <summary> Returns the content type of this cell
			/// 
			/// </summary>
			/// <returns> the content type for this cell
			/// </returns>
			CellType Type
			{
				get;
				
			}
			/// <summary> Indicates whether or not this cell is hidden, by virtue of either
			/// the entire row or column being collapsed
			/// 
			/// </summary>
			/// <returns> TRUE if this cell is hidden, FALSE otherwise
			/// </returns>
			bool Hidden
			{
				get;
				
			}
			/// <summary> Quick and dirty function to return the contents of this cell as a string.
			/// For more complex manipulation of the contents, it is necessary to cast
			/// this interface to correct subinterface
			/// 
			/// </summary>
			/// <returns> the contents of this cell as a string
			/// </returns>
			string Contents
			{
				get;
				
			}


			/// <summary> Gets the value for this cell.
			/// Depending on the cell, Value can be a string, DateTime, double.
			/// </summary>
			/// <returns> the cell value
			/// </returns>
			object Value
			{
				get;
					
			}

			/// <summary> Gets the cell format which applies to this cell
			/// Note that for cell with a cell type of EMPTY, which has no formatting
			/// information, this method will return null.  Some empty cells (eg. on
			/// template spreadsheets) may have a cell type of EMPTY, but will
			/// actually contain formatting information
			/// 
			/// </summary>
			/// <returns> the cell format applied to this cell, or NULL if this is an
			/// empty cell
			/// </returns>
			NExcel.Format.CellFormat CellFormat
			{
				get;
				
			}
		}
}
