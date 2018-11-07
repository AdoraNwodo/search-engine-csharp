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
using NExcel.Biff.Formula;
namespace NExcel.Biff
{
	
	/// <summary> Implementation class for the Range interface.  This merely
	/// holds the raw range information, and when the time comes, it
	/// interrogates the workbook for the object.
	/// This does not keep handles to the objects for performance reasons,
	/// as this could impact garbage collection on larger spreadsheets
	/// </summary>
	public class RangeImpl : Range
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
				Sheet s = workbook.getReadSheet(sheet1);
				
				if (column1 < s.Columns && row1 < s.Rows)
				{
					return s.getCell(column1, row1);
				}
				else
				{
					return new EmptyCell(column1, row1);
				}
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
				Sheet s = workbook.getReadSheet(sheet2);
				
				if (column2 < s.Columns && row2 < s.Rows)
				{
					return s.getCell(column2, row2);
				}
				else
				{
					return new EmptyCell(column2, row2);
				}
			}
			
		}
		/// <summary> Gets the index of the first sheet in the range
		/// 
		/// </summary>
		/// <returns> the index of the first sheet in the range
		/// </returns>
		virtual public int FirstSheetIndex
		{
			get
			{
				return sheet1;
			}
			
		}
		/// <summary> Gets the index of the last sheet in the range
		/// 
		/// </summary>
		/// <returns> the index of the last sheet in the range
		/// </returns>
		virtual public int LastSheetIndex
		{
			get
			{
				return sheet2;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> A handle to the workbook</summary>
		private WorkbookMethods workbook;
		
		/// <summary> The sheet index containing the column at the top left</summary>
		private int sheet1;
		
		/// <summary> The column number of the cell at the top left of the range</summary>
		private int column1;
		
		/// <summary> The row number of the cell at the top left of the range</summary>
		private int row1;
		
		/// <summary> The sheet index of the cell at the bottom right</summary>
		private int sheet2;
		
		/// <summary> The column index of the cell at the bottom right</summary>
		private int column2;
		
		/// <summary> The row index of the cell at the bottom right</summary>
		private int row2;
		
		/// <summary> Constructor</summary>
		/// <param name="w">the workbook
		/// </param>
		/// <param name="es">the external sheet
		/// </param>
		/// <param name="s1">the sheet of the top left cell of the range
		/// </param>
		/// <param name="c1">the column number of the top left cell of the range
		/// </param>
		/// <param name="r1">the row number of the top left cell of the range
		/// </param>
		/// <param name="s2">the sheet of the bottom right cell
		/// </param>
		/// <param name="c2">the column number of the bottom right cell of the range
		/// </param>
		/// <param name="r2">the row number of the bottomr right cell of the range
		/// </param>
		public RangeImpl(WorkbookMethods w, int s1, int c1, int r1, int s2, int c2, int r2)
		{
			workbook = w;
			sheet1 = s1;
			sheet2 = s2;
			row1 = r1;
			row2 = r2;
			column1 = c1;
			column2 = c2;
		}
		static RangeImpl()
		{
			logger = Logger.getLogger(typeof(RangeImpl));
		}
	}
}
