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
using common;
using NExcel;
using NExcel.Format;
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	/// <summary> Abstract class for all records which actually contain cell values</summary>
	public abstract class CellValue:RecordData, Cell
	{
		/// <summary> Interface method which returns the row number of this cell
		/// 
		/// </summary>
		/// <returns> the zero base row number
		/// </returns>
		virtual public int Row
		{
			get
			{
				return row;
			}
			
		}
		/// <summary> Interface method which returns the column number of this cell
		/// 
		/// </summary>
		/// <returns> the zero based column number
		/// </returns>
		virtual public int Column
		{
			get
			{
				return column;
			}
			
		}
		/// <summary> Gets the XFRecord corresponding to the index number.  Used when
		/// copying a spreadsheet
		/// 
		/// </summary>
		/// <returns> the xf index for this cell
		/// </returns>
		virtual public int XFIndex
		{
			get
			{
				return xfIndex;
			}
			
		}
		/// <summary> Gets the CellFormat object for this cell.  Used by the WritableWorkbook
		/// API
		/// 
		/// </summary>
		/// <returns> the CellFormat used for this cell
		/// </returns>
		virtual public NExcel.Format.CellFormat CellFormat
		{
			get
			{
				if (!initialized)
				{
					format = formattingRecords.getXFRecord(xfIndex);
					initialized = true;
				}
				
				return format;
			}
		
		}

		virtual public CellType Type
		{
			get
			{
				return null;
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
				
				if (cir != null && (cir.Width == 0 || cir.Hidden))
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
		/// <summary> Accessor for the sheet
		/// 
		/// </summary>
		/// <returns> the sheet
		/// </returns>
		virtual protected internal SheetImpl Sheet
		{
			get
			{
				return sheet;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The row number of this cell record</summary>
		private int row;
		
		/// <summary> The column number of this cell record</summary>
		private int column;
		
		/// <summary> The XF index</summary>
		private int xfIndex;
		
		/// <summary> A handle to the formatting records, so that we can
		/// retrieve the formatting information
		/// </summary>
		private FormattingRecords formattingRecords;
		
		/// <summary> A lazy initialize flag for the cell format</summary>
		private bool initialized;
		
		/// <summary> The cell format</summary>
		private XFRecord format;
		
		/// <summary> A handle back to the sheet</summary>
		private SheetImpl sheet;
		
		/// <summary> Constructs this object from the raw cell data
		/// 
		/// </summary>
		/// <param name="t">the raw cell data
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="si">the sheet containing this cell
		/// </param>
		protected internal CellValue(Record t, FormattingRecords fr, SheetImpl si):base(t)
		{
			sbyte[] data = getRecord().Data;
			row = IntegerHelper.getInt(data[0], data[1]);
			column = IntegerHelper.getInt(data[2], data[3]);
			xfIndex = IntegerHelper.getInt(data[4], data[5]);
			sheet = si;
			formattingRecords = fr;
			initialized = false;
		}
		

		virtual public string Contents
		{
			get
			{
				return null;
			}
		}
	
		/// <summary>
		/// Returns a empty value.
		/// </summary>
		virtual public object Value
		{
			get
			{
				return null;
			}
		}


		static CellValue()
		{
			logger = Logger.getLogger(typeof(CellValue));
		}
	}
}
