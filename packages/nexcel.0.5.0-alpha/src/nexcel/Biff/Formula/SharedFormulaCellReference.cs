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
	class SharedFormulaCellReference:Operand, ParsedThing
	{
		virtual public int Column
		{
			get
			{
				return column;
			}
			
		}
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
				data[0] = Token.REF.Code;
				
				IntegerHelper.getTwoBytes(row, data, 1);
				
				int columnMask = column;
				
				if (columnRelative)
				{
					columnMask |= 0x4000;
				}
				
				if (rowRelative)
				{
					columnMask |= 0x8000;
				}
				
				IntegerHelper.getTwoBytes(columnMask, data, 3);
				
				return data;
			}
			
		}
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
		/// <param name="the">cell the formula is relative to
		/// </param>
		public SharedFormulaCellReference(Cell rt)
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
			row = IntegerHelper.getShort(data[pos], data[pos + 1]);
			
			int columnMask = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
			
			column = (sbyte) (columnMask & 0xff);
			columnRelative = ((columnMask & 0x4000) != 0);
			rowRelative = ((columnMask & 0x8000) != 0);
			
			if (columnRelative)
			{
				column = relativeTo.Column + column;
			}
			
			if (rowRelative)
			{
				row = relativeTo.Row + row;
			}
			
			return 4;
		}
		
		public override void  getString(System.Text.StringBuilder buf)
		{
			CellReferenceHelper.getCellReference(column, row, buf);
		}
	}
}
