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
using NExcel.Format;
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	/// <summary> A blank cell value, initialized indirectly from a multiple biff record
	/// rather than directly from the binary data
	/// </summary>
	class MulBlankCell : Cell
	{
		/// <summary> Accessor for the row
		/// 
		/// </summary>
		/// <returns> the zero based row
		/// </returns>
		virtual public int Row
		{
			get
			{
				return row;
			}
			
		}
		/// <summary> Accessor for the column
		/// 
		/// </summary>
		/// <returns> the zero based column
		/// </returns>
		virtual public int Column
		{
			get
			{
				return column;
			}
			
		}
		/// <summary> Accessor for the contents as a string
		/// 
		/// </summary>
		/// <returns> the value as a string
		/// </returns>
		virtual public string Contents
		{
			get
			{
				return "";
			}
			
		}

		/// <summary>
		/// Returns the value.
		/// </summary>
		virtual public object Value
		{
			get
			{
				return "";
			}
		}

		/// <summary> Accessor for the cell type
		/// 
		/// </summary>
		/// <returns> the cell type
		/// </returns>
		virtual public CellType Type
		{
			get
			{
				return CellType.EMPTY;
			}
			
		}
		/// <summary> Gets the cell format for this cell
		/// 
		/// </summary>
		/// <returns>  the cell format for these cells
		/// </returns>
		virtual public NExcel.Format.CellFormat CellFormat
		{
			get
			{
				if (!initialized)
				{
					cellFormat = formattingRecords.getXFRecord(xfIndex);
					initialized = true;
				}
				
				return cellFormat;
			}
			
		}
		/// <summary> Determines whether or not this cell has been hidden
		/// 
		/// </summary>
		/// <returns> TRUE if this cell has been hidden, FALSE otherwise
		/// </returns>
		virtual public bool Hidden
		{
			get
			{
				ColumnInfoRecord cir = sheet.getColumnInfo(column);
				
				if (cir != null && cir.Width == 0)
				{
					return true;
				}
				
				RowRecord rr = sheet.getRowInfo(row);
				
				if (rr != null && (rr.RowHeight == 0 || rr.isCollapsed()))
				{
					return true;
				}
				
				return false;
			}
			
		}
		/// <summary> The row containing this blank</summary>
		private int row;
		/// <summary> The column containing this blank</summary>
		private int column;
		/// <summary> The raw cell format</summary>
		private NExcel.Format.CellFormat cellFormat;
		
		/// <summary> The index to the XF Record</summary>
		private int xfIndex;
		
		/// <summary> A handle to the formatting records</summary>
		private FormattingRecords formattingRecords;
		
		/// <summary> A flag to indicate whether this object's formatting things have
		/// been initialized
		/// </summary>
		private bool initialized;
		
		/// <summary> A handle to the sheet</summary>
		private SheetImpl sheet;
		
		/// <summary> Constructs this cell
		/// 
		/// </summary>
		/// <param name="r">the zero based row
		/// </param>
		/// <param name="c">the zero base column
		/// </param>
		/// <param name="xfi">the xf index
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		public MulBlankCell(int r, int c, int xfi, FormattingRecords fr, SheetImpl si)
		{
			row = r;
			column = c;
			xfIndex = xfi;
			formattingRecords = fr;
			sheet = si;
			initialized = false;
		}
	}
}
