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
namespace NExcel.Biff
{
	
	/// <summary> Implementation class for the Range interface.  This merely
	/// holds the raw range information.  This implementation is used
	/// for ranges which are present on the current working sheet, so the
	/// getSheetIndex merely returns -1
	/// </summary>
	public class SheetRangeImpl : Range
	{
		/// <summary> Gets the cell at the top left of this range
		/// 
		/// </summary>
		/// <returns> the cell at the top left
		/// </returns>
		virtual public Cell TopLeft
		{
			get
			{
				return sheet.getCell(column1, row1);
			}
			
		}
		/// <summary> Gets the cell at the bottom right of this range
		/// 
		/// </summary>
		/// <returns> the cell at the bottom right
		/// </returns>
		virtual public Cell BottomRight
		{
			get
			{
				return sheet.getCell(column2, row2);
			}
			
		}
		/// <summary> Not supported.  Returns -1, indicating that it refers to the current
		/// sheet
		/// 
		/// </summary>
		/// <returns> -1
		/// </returns>
		virtual public int FirstSheetIndex
		{
			get
			{
				return - 1;
			}
			
		}
		/// <summary> Not supported.  Returns -1, indicating that it refers to the current
		/// sheet
		/// 
		/// </summary>
		/// <returns> -1
		/// </returns>
		virtual public int LastSheetIndex
		{
			get
			{
				return - 1;
			}
			
		}
		/// <summary> A handle to the sheet containing this range</summary>
		private Sheet sheet;
		
		/// <summary> The column number of the cell at the top left of the range</summary>
		private int column1;
		
		/// <summary> The row number of the cell at the top left of the range</summary>
		private int row1;
		
		/// <summary> The column index of the cell at the bottom right</summary>
		private int column2;
		
		/// <summary> The row index of the cell at the bottom right</summary>
		private int row2;
		
		/// <summary> Constructor</summary>
		/// <param name="s">the sheet containing the range
		/// </param>
		/// <param name="c1">the column number of the top left cell of the range
		/// </param>
		/// <param name="r1">the row number of the top left cell of the range
		/// </param>
		/// <param name="c2">the column number of the bottom right cell of the range
		/// </param>
		/// <param name="r2">the row number of the bottomr right cell of the range
		/// </param>
		public SheetRangeImpl(Sheet s, int c1, int r1, int c2, int r2)
		{
			sheet = s;
			row1 = r1;
			row2 = r2;
			column1 = c1;
			column2 = c2;
		}
		
		/// <summary> A copy constructor used for copying ranges between sheets
		/// 
		/// </summary>
		/// <param name="c">the range to copy from
		/// </param>
		/// <param name="s">the writable sheet
		/// </param>
		public SheetRangeImpl(SheetRangeImpl c, Sheet s)
		{
			sheet = s;
			row1 = c.row1;
			row2 = c.row2;
			column1 = c.column1;
			column2 = c.column2;
		}
		
		/// <summary> Sees whether there are any intersections between this range and the
		/// range passed in.  This method is used internally by the WritableSheet to
		/// verify the integrity of merged cells, hyperlinks etc.  Ranges are
		/// only ever compared for the same sheet
		/// 
		/// </summary>
		/// <param name="range">the range to compare against
		/// </param>
		/// <returns> TRUE if the ranges intersect, FALSE otherwise
		/// </returns>
		public virtual bool intersects(SheetRangeImpl range)
		{
			if (range == this)
			{
				return true;
			}
			
			if (row2 < range.row1 || row1 > range.row2 || column2 < range.column1 || column1 > range.column2)
			{
				return false;
			}
			
			return true;
		}
		
		/// <summary> To string method - primarily used during debugging
		/// 
		/// </summary>
		/// <returns> the string version of this object
		/// </returns>
		public override string ToString()
		{
			System.Text.StringBuilder sb = new System.Text.StringBuilder();
			CellReferenceHelper.getCellReference(column1, row1, sb);
			sb.Append('-');
			CellReferenceHelper.getCellReference(column2, row2, sb);
			return sb.ToString();
		}
		
		/// <summary> A row has been inserted, so adjust the range objects accordingly
		/// 
		/// </summary>
		/// <param name="r">the row which has been inserted
		/// </param>
		public virtual void  insertRow(int r)
		{
			if (r > row2)
			{
				return ;
			}
			
			if (r <= row1)
			{
				row1++;
			}
			
			if (r <= row2)
			{
				row2++;
			}
		}
		
		/// <summary> A column has been inserted, so adjust the range objects accordingly
		/// 
		/// </summary>
		/// <param name="c">the column which has been inserted
		/// </param>
		public virtual void  insertColumn(int c)
		{
			if (c > column2)
			{
				return ;
			}
			
			if (c <= column1)
			{
				column1++;
			}
			
			if (c <= column2)
			{
				column2++;
			}
		}
		
		/// <summary> A row has been removed, so adjust the range objects accordingly
		/// 
		/// </summary>
		/// <param name="r">the row which has been inserted
		/// </param>
		public virtual void  removeRow(int r)
		{
			if (r > row2)
			{
				return ;
			}
			
			if (r < row1)
			{
				row1--;
			}
			
			if (r < row2)
			{
				row2--;
			}
		}
		
		/// <summary> A column has been removed, so adjust the range objects accordingly
		/// 
		/// </summary>
		/// <param name="c">the column which has been removed
		/// </param>
		public virtual void  removeColumn(int c)
		{
			if (c > column2)
			{
				return ;
			}
			
			if (c < column1)
			{
				column1--;
			}
			
			if (c < column2)
			{
				column2--;
			}
		}
		
		/// <summary> Standard hash code method
		/// 
		/// </summary>
		/// <returns> the hash code
		/// </returns>
		public override int GetHashCode()
		{
			return 0xffff ^ row1 ^ row2 ^ column1 ^ column2;
		}
		
		/// <summary> Standard equals method
		/// 
		/// </summary>
		/// <param name="o">the object to compare
		/// </param>
		/// <returns> TRUE if the two objects are the same, FALSE otherwise
		/// </returns>
		public  override bool Equals(System.Object o)
		{
			if (o == this)
			{
				return true;
			}
			
			if (!(o is SheetRangeImpl))
			{
				return false;
			}
			
			SheetRangeImpl compare = (SheetRangeImpl) o;
			
			return (column1 == compare.column1 && column2 == compare.column2 && row1 == compare.row1 && row2 == compare.row2);
		}
	}
}
