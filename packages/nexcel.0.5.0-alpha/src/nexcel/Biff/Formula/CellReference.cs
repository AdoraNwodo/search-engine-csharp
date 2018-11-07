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
using NExcel.Biff;
namespace NExcel.Biff.Formula
{
	
	/// <summary> A cell reference in a formula</summary>
	class CellReference:Operand, ParsedThing
	{
		/// <summary> Accessor for the column
		/// 
		/// </summary>
		/// <returns> the column
		/// </returns>
		virtual public int Column
		{
			get
			{
				return column;
			}
			
		}
		/// <summary> Accessor for the row
		/// 
		/// </summary>
		/// <returns> the row
		/// </returns>
		virtual public int Row
		{
			get
			{
				return row;
			}
			
		}
		/// <summary> Gets the token representation of this item in RPN
		/// 
		/// </summary>
		/// <returns> the bytes applicable to this formula
		/// </returns>
		override internal sbyte[] Bytes
		{
			get
			{
				sbyte[] data = new sbyte[5];
				data[0] = !useAlternateCode()?Token.REF.Code:Token.REF.Code2;
				
				IntegerHelper.getTwoBytes(row, data, 1);
				
				int grcol = column;
				
				// Set the row/column relative bits if applicable
				if (rowRelative)
				{
					grcol |= 0x8000;
				}
				
				if (columnRelative)
				{
					grcol |= 0x4000;
				}
				
				IntegerHelper.getTwoBytes(grcol, data, 3);
				
				return data;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> Indicates whether the column reference is relative or absolute</summary>
		private bool columnRelative;
		
		/// <summary> Indicates whether the row reference is relative or absolute</summary>
		private bool rowRelative;
		
		/// <summary> The column reference</summary>
		private int column;
		
		/// <summary> The row reference</summary>
		private int row;
		
		/// <summary> The cell containing the formula.  Stored in order to determine
		/// relative cell values
		/// </summary>
		private Cell relativeTo;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="rt">the cell containing the formula
		/// </param>
		public CellReference(Cell rt)
		{
			relativeTo = rt;
		}
		
		/// <summary> Constructor</summary>
		public CellReference()
		{
		}
		
		/// <summary> Constructor invoked when parsing a text string
		/// 
		/// </summary>
		/// <param name="s">the string being parsed
		/// </param>
		public CellReference(string s)
		{
			column = CellReferenceHelper.getColumn(s);
			row = CellReferenceHelper.getRow(s);
			columnRelative = CellReferenceHelper.isColumnRelative(s);
			rowRelative = CellReferenceHelper.isRowRelative(s);
		}
		
		/// <summary> Reads the ptg data from the array starting at the specified position
		/// 
		/// </summary>
		/// <param name="data">the RPN array
		/// </param>
		/// <param name="pos">the current position in the array, excluding the ptg identifier
		/// </param>
		/// <returns> the number of bytes read
		/// </returns>
		public virtual int read(sbyte[] data, int pos)
		{
			row = IntegerHelper.getInt(data[pos], data[pos + 1]);
			int columnMask = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
			column = columnMask & 0x00ff;
			columnRelative = ((columnMask & 0x4000) != 0);
			rowRelative = ((columnMask & 0x8000) != 0);
			
			return 4;
		}
		
		/// <summary> Gets the cell reference as a string for this item
		/// 
		/// </summary>
		/// <param name="buf">the string buffer to populate
		/// </param>
		public override void  getString(System.Text.StringBuilder buf)
		{
			CellReferenceHelper.getCellReference(column, !columnRelative, row, !rowRelative, buf);
		}
		
		/// <summary> Adjusts all the relative cell references in this formula by the
		/// amount specified.  Used when copying formulas
		/// 
		/// </summary>
		/// <param name="colAdjust">the amount to add on to each relative cell reference
		/// </param>
		/// <param name="rowAdjust">the amount to add on to each relative row reference
		/// </param>
		public override void  adjustRelativeCellReferences(int colAdjust, int rowAdjust)
		{
			if (columnRelative)
			{
				column += colAdjust;
			}
			
			if (rowRelative)
			{
				row += rowAdjust;
			}
		}
		
		/// <summary> Called when a column is inserted on the specified sheet.  Tells
		/// the formula  parser to update all of its cell references beyond this
		/// column
		/// 
		/// </summary>
		/// <param name="sheetIndex">the sheet on which the column was inserted
		/// </param>
		/// <param name="col">the column number which was inserted
		/// </param>
		/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
		/// column was inserted, FALSE otherwise
		/// </param>
		public override void  columnInserted(int sheetIndex, int col, bool currentSheet)
		{
			if (!currentSheet)
			{
				return ;
			}
			
			if (column >= col)
			{
				column++;
			}
		}
		
		/// <summary> Called when a column is inserted on the specified sheet.  Tells
		/// the formula  parser to update all of its cell references beyond this
		/// column
		/// 
		/// </summary>
		/// <param name="sheetIndex">the sheet on which the column was removed
		/// </param>
		/// <param name="col">the column number which was removed
		/// </param>
		/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
		/// column was inserted, FALSE otherwise
		/// </param>
		internal override void  columnRemoved(int sheetIndex, int col, bool currentSheet)
		{
			if (!currentSheet)
			{
				return ;
			}
			
			if (column >= col)
			{
				column--;
			}
		}
		
		/// <summary> Called when a column is inserted on the specified sheet.  Tells
		/// the formula  parser to update all of its cell references beyond this
		/// column
		/// 
		/// </summary>
		/// <param name="sheetIndex">the sheet on which the row was inserted
		/// </param>
		/// <param name="row">the row number which was inserted
		/// </param>
		/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
		/// column was inserted, FALSE otherwise
		/// </param>
		internal override void  rowInserted(int sheetIndex, int r, bool currentSheet)
		{
			if (!currentSheet)
			{
				return ;
			}
			
			if (row >= r)
			{
				row++;
			}
		}
		
		/// <summary> Called when a column is inserted on the specified sheet.  Tells
		/// the formula  parser to update all of its cell references beyond this
		/// column
		/// 
		/// </summary>
		/// <param name="sheetIndex">the sheet on which the row was removed
		/// </param>
		/// <param name="row">the row number which was removed
		/// </param>
		/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
		/// column was inserted, FALSE otherwise
		/// </param>
		internal override void  rowRemoved(int sheetIndex, int r, bool currentSheet)
		{
			if (!currentSheet)
			{
				return ;
			}
			
			if (row >= r)
			{
				row--;
			}
		}
		static CellReference()
		{
			logger = Logger.getLogger(typeof(CellReference));
		}
	}
}
