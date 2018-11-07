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
using CellFormat = NExcel.Format.CellFormat;
namespace NExcel
{
	
	/// <summary> This is a bean which client applications may use to get/set various
	/// properties for a row or column on a spreadsheet
	/// </summary>
	public sealed class CellView
	{
		/// <summary> Accessor for the hidden nature of this row/column
		/// 
		/// </summary>
		/// <returns> TRUE if this row/column is hidden, FALSE otherwise
		/// </returns>
		/// <summary> Sets the hidden status of this row/column
		/// 
		/// </summary>
		/// <param name="h">the hidden flag
		/// </param>
		public bool Hidden
		{
			get
			{
				return hidden;
			}
			
			set
			{
				hidden = value;
			}
			
		}
		/// <summary> Gets the width of the column in characters or the height of the
		/// row in 1/20ths
		/// 
		/// </summary>
		/// <returns> the dimension
		/// </returns>
		/// <deprecated> use getSize() instead
		/// </deprecated>
		/// <summary> Sets the dimension for this view
		/// 
		/// </summary>
		/// <param name="d">the width of the column in characters, or the height of the
		/// row in 1/20ths of a point
		/// </param>
		/// <deprecated> use the setSize method instead
		/// </deprecated>
		public int Dimension
		{
			get
			{
				return dimension;
			}
			
			set
			{
				dimension = value;
				depUsed_Renamed_Field = true;
			}
			
		}
		/// <summary> Gets the width of the column in characters multiplied by 256, or the 
		/// height of the row in 1/20ths of a point
		/// 
		/// </summary>
		/// <returns> the dimension
		/// </returns>
		/// <summary> Sets the dimension for this view
		/// 
		/// </summary>
		/// <param name="d">the width of the column in characters multiplied by 256, 
		/// or the height of the
		/// row in 1/20ths of a point
		/// </param>
		public int Size
		{
			get
			{
				return size;
			}
			
			set
			{
				size = value;
				depUsed_Renamed_Field = false;
			}
			
		}
		/// <summary> Accessor for the cell format for this group.
		/// 
		/// </summary>
		/// <returns> the format for the column/row, or NULL if no format was
		/// specified
		/// </returns>
		/// <summary> Sets the cell format for this group of cells
		/// 
		/// </summary>
		/// <param name="cf">the format for every cell in the column/row
		/// </param>
		public NExcel.Format.CellFormat Format
		{
			get
			{
				return format;
			}
			
			set
			{
				format = value;
			}
			
		}
		/// <summary> The dimension for the associated group of cells.  For columns this
		/// will be width in characters, for rows this will be the
		/// height in points
		/// This attribute is deprecated in favour of the size attribute
		/// </summary>
		private int dimension;
		
		/// <summary> The size for the associated group of cells.  For columns this
		/// will be width in characters multiplied by 256, for rows this will be the
		/// height in points
		/// </summary>
		private int size;
		
		/// <summary> Indicates whether the deprecated function was used to set the dimension</summary>
		private bool depUsed_Renamed_Field;
		
		/// <summary> Indicates whether or not this sheet is hidden</summary>
		private bool hidden;
		
		/// <summary> The cell format for the row/column</summary>
		private NExcel.Format.CellFormat format;
		
		/// <summary> Default constructor</summary>
		public CellView()
		{
			hidden = false;
			depUsed_Renamed_Field = false;
			dimension = 1;
			size = 1;
		}
		
		/// <summary> Accessor for the depUsed attribute
		/// 
		/// </summary>
		/// <returns> TRUE if the deprecated methods were used to set the size,
		/// FALSE otherwise
		/// </returns>
		public bool depUsed()
		{
			return depUsed_Renamed_Field;
		}
	}
}
