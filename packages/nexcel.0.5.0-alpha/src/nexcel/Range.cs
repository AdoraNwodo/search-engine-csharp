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
namespace NExcel
{
	
	/// <summary> Represents a 3-D range of cells in a workbook.  This object is
	/// returned by the method findByName in a workbook
	/// </summary>
	public interface Range
		{
			/// <summary> Gets the cell at the top left of this range
			/// 
			/// </summary>
			/// <returns> the cell at the top left
			/// </returns>
			Cell TopLeft
			{
				get;
				
			}
			/// <summary> Gets the cell at the bottom right of this range
			/// 
			/// </summary>
			/// <returns> the cell at the bottom right
			/// </returns>
			Cell BottomRight
			{
				get;
				
			}
			/// <summary> Gets the index of the first sheet in the range
			/// 
			/// </summary>
			/// <returns> the index of the first sheet in the range
			/// </returns>
			int FirstSheetIndex
			{
				get;
				
			}
			/// <summary> Gets the index of the last sheet in the range
			/// 
			/// </summary>
			/// <returns> the index of the last sheet in the range
			/// </returns>
			int LastSheetIndex
			{
				get;
				
			}
		}
}
