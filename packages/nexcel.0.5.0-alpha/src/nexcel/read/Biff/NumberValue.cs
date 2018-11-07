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
using NExcelUtils;
using NExcel;
using NExcel.Format;
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	/// <summary> A numerical cell value, initialized indirectly from a multiple biff record
	/// rather than directly from the binary data
	/// </summary>
	class NumberValue : NumberCell
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
		/// <summary> Accessor for the value
		/// 
		/// </summary>
		/// <returns> the value
		/// </returns>
		virtual public double DoubleValue
		{
			get
			{
			return _Value;
			}
		}

		/// <summary>
		/// Returns the value.
		/// </summary>
		public virtual object Value
		{
			get
			{
				return this._Value;
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
				return string.Format(format, "{0}", _Value);
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
				return CellType.NUMBER;
			}
			
		}
		/// <summary> Gets the cell format
		/// 
		/// </summary>
		/// <returns> the cell format
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
		/// <summary> The row containing this number</summary>
		private int row;
		/// <summary> The column containing this number</summary>
		private int column;
		/// <summary> The value of this number</summary>
		private double _Value;
		
		/// <summary> The cell format</summary>
		private NumberFormatInfo format;
		
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
		
		/// <summary> The format in which to return this number as a string</summary>
		private static NumberFormatInfo defaultFormat;
		
		/// <summary> Constructs this number
		/// 
		/// </summary>
		/// <param name="r">the zero based row
		/// </param>
		/// <param name="c">the zero base column
		/// </param>
		/// <param name="val">the value
		/// </param>
		/// <param name="xfi">the xf index
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		public NumberValue(int r, int c, double val, int xfi, FormattingRecords fr, SheetImpl si)
		{
			row = r;
			column = c;
			_Value = val;
			format = defaultFormat;
			xfIndex = xfi;
			formattingRecords = fr;
			sheet = si;
			initialized = false;
		}
		
		/// <summary> Sets the format for the number based on the Excel spreadsheets' format.
		/// This is called from SheetImpl when it has been definitely established
		/// that this cell is a number and not a date
		/// 
		/// </summary>
		/// <param name="f">the format
		/// </param>
		internal void  setNumberFormat(NumberFormatInfo f)
		{
			if (f != null)
			{
				format = f;
			}
		}
		
		/// <summary> Gets the NumberFormatInfo used to format this cell.  This is the java
		/// equivalent of the Excel format
		/// 
		/// </summary>
		/// <returns> the NumberFormatInfo used to format the cell
		/// </returns>
		public virtual NumberFormatInfo NumberFormat
		{
			get
			{
				return format;
			}
		}

		static NumberValue()
		{
			defaultFormat = new NumberFormatInfo("#.###");
		}
	}
}
