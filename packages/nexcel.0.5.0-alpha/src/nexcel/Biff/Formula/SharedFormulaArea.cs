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
using NExcel.Biff;
namespace NExcel.Biff.Formula
{
	
	/// <summary> A cell reference in a formula</summary>
	class SharedFormulaArea:Operand, ParsedThing
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
				data[0] = Token.AREA.Code;
				
				// Use absolute references for columns, so don't bother about
				// the col relative/row relative bits
				IntegerHelper.getTwoBytes(rowFirst, data, 1);
				IntegerHelper.getTwoBytes(rowLast, data, 3);
				IntegerHelper.getTwoBytes(columnFirst, data, 5);
				IntegerHelper.getTwoBytes(columnLast, data, 7);
				
				return data;
			}
			
		}
		private int columnFirst;
		private int rowFirst;
		private int columnLast;
		private int rowLast;
		
		private bool columnFirstRelative;
		private bool rowFirstRelative;
		private bool columnLastRelative;
		private bool rowLastRelative;
		
		/// <summary> The cell containing the formula.  Stored in order to determine
		/// relative cell values
		/// </summary>
		private Cell relativeTo;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="the">cell the formula is relative to
		/// </param>
		public SharedFormulaArea(Cell rt)
		{
			relativeTo = rt;
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
			// Preserve signage on column and row values, because they will
			// probably be relative
			
			rowFirst = IntegerHelper.getShort(data[pos], data[pos + 1]);
			rowLast = IntegerHelper.getShort(data[pos + 2], data[pos + 3]);
			
			int columnMask = IntegerHelper.getInt(data[pos + 4], data[pos + 5]);
			columnFirst = columnMask & 0x00ff;
			columnFirstRelative = ((columnMask & 0x4000) != 0);
			rowFirstRelative = ((columnMask & 0x8000) != 0);
			
			if (columnFirstRelative)
			{
				columnFirst = relativeTo.Column + columnFirst;
			}
			
			if (rowFirstRelative)
			{
				rowFirst = relativeTo.Row + rowFirst;
			}
			
			columnMask = IntegerHelper.getInt(data[pos + 6], data[pos + 7]);
			columnLast = columnMask & 0x00ff;
			
			columnLastRelative = ((columnMask & 0x4000) != 0);
			rowLastRelative = ((columnMask & 0x8000) != 0);
			
			if (columnLastRelative)
			{
				columnLast = relativeTo.Column + columnLast;
			}
			
			if (rowLastRelative)
			{
				rowLast = relativeTo.Row + rowLast;
			}
			
			
			return 8;
		}
		
		public override void  getString(System.Text.StringBuilder buf)
		{
			CellReferenceHelper.getCellReference(columnFirst, rowFirst, buf);
			buf.Append(':');
			CellReferenceHelper.getCellReference(columnLast, rowLast, buf);
		}
	}
}
