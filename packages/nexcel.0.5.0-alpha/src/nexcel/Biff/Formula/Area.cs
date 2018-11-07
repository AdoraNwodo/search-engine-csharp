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
using System.Collections;
using common;
using NExcel.Biff;
namespace NExcel.Biff.Formula
{
	
	/// <summary> A nested class to hold range information</summary>
	class Area:Operand, ParsedThing
	{
		virtual internal int FirstColumn
		{
			get
			{
				return columnFirst;
			}
			
		}
		virtual internal int FirstRow
		{
			get
			{
				return rowFirst;
			}
			
		}
		virtual internal int LastColumn
		{
			get
			{
				return columnLast;
			}
			
		}
		virtual internal int LastRow
		{
			get
			{
				return rowLast;
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
				sbyte[] data = new sbyte[9];
				data[0] = !useAlternateCode()?Token.AREA.Code:Token.AREA.Code2;
				
				IntegerHelper.getTwoBytes(rowFirst, data, 1);
				IntegerHelper.getTwoBytes(rowLast, data, 3);
				
				int grcol = columnFirst;
				
				// Set the row/column relative bits if applicable
				if (rowFirstRelative)
				{
					grcol |= 0x8000;
				}
				
				if (columnFirstRelative)
				{
					grcol |= 0x4000;
				}
				
				IntegerHelper.getTwoBytes(grcol, data, 5);
				
				grcol = columnLast;
				
				// Set the row/column relative bits if applicable
				if (rowLastRelative)
				{
					grcol |= 0x8000;
				}
				
				if (columnLastRelative)
				{
					grcol |= 0x4000;
				}
				
				IntegerHelper.getTwoBytes(grcol, data, 7);
				
				return data;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		private int columnFirst;
		private int rowFirst;
		private int columnLast;
		private int rowLast;
		private bool columnFirstRelative;
		private bool rowFirstRelative;
		private bool columnLastRelative;
		private bool rowLastRelative;
		
		/// <summary> Constructor</summary>
		internal Area()
		{
		}
		
		/// <summary> Constructor invoked when parsing a string formula
		/// 
		/// </summary>
		/// <param name="s">the string to parse
		/// </param>
		internal Area(string s)
		{
			int seppos = s.IndexOf(":");
			Assert.verify(seppos != - 1);
			string startcell = s.Substring(0, (seppos) - (0));
			string endcell = s.Substring(seppos + 1);
			
			columnFirst = CellReferenceHelper.getColumn(startcell);
			rowFirst = CellReferenceHelper.getRow(startcell);
			columnLast = CellReferenceHelper.getColumn(endcell);
			rowLast = CellReferenceHelper.getRow(endcell);
			
			columnFirstRelative = CellReferenceHelper.isColumnRelative(startcell);
			rowFirstRelative = CellReferenceHelper.isRowRelative(startcell);
			columnLastRelative = CellReferenceHelper.isColumnRelative(endcell);
			rowLastRelative = CellReferenceHelper.isRowRelative(endcell);
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
			rowFirst = IntegerHelper.getInt(data[pos], data[pos + 1]);
			rowLast = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
			int columnMask = IntegerHelper.getInt(data[pos + 4], data[pos + 5]);
			columnFirst = columnMask & 0x00ff;
			columnFirstRelative = ((columnMask & 0x4000) != 0);
			rowFirstRelative = ((columnMask & 0x8000) != 0);
			columnMask = IntegerHelper.getInt(data[pos + 6], data[pos + 7]);
			columnLast = columnMask & 0x00ff;
			columnLastRelative = ((columnMask & 0x4000) != 0);
			rowLastRelative = ((columnMask & 0x8000) != 0);
			
			return 8;
		}
		
		/// <summary> Gets the string representation of this item
		/// 
		/// </summary>
		/// <param name="">buf
		/// </param>
		public override void  getString(System.Text.StringBuilder buf)
		{
			CellReferenceHelper.getCellReference(columnFirst, rowFirst, buf);
			buf.Append(':');
			CellReferenceHelper.getCellReference(columnLast, rowLast, buf);
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
			if (columnFirstRelative)
			{
				columnFirst += colAdjust;
			}
			
			if (columnLastRelative)
			{
				columnLast += colAdjust;
			}
			
			if (rowFirstRelative)
			{
				rowFirst += rowAdjust;
			}
			
			if (rowLastRelative)
			{
				rowLast += rowAdjust;
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
			
			if (col <= columnFirst)
			{
				columnFirst++;
			}
			
			if (col <= columnLast)
			{
				columnLast++;
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
			
			if (col < columnFirst)
			{
				columnFirst--;
			}
			
			if (col <= columnLast)
			{
				columnLast--;
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
		internal override void  rowInserted(int sheetIndex, int row, bool currentSheet)
		{
			if (!currentSheet)
			{
				return ;
			}
			
			if (row <= rowFirst)
			{
				rowFirst++;
			}
			
			if (row <= rowLast)
			{
				rowLast++;
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
		internal override void  rowRemoved(int sheetIndex, int row, bool currentSheet)
		{
			if (!currentSheet)
			{
				return ;
			}
			
			if (row < rowFirst)
			{
				rowFirst--;
			}
			
			if (row <= rowLast)
			{
				rowLast--;
			}
		}
		static Area()
		{
			logger = Logger.getLogger(typeof(Area));
		}
	}
}
