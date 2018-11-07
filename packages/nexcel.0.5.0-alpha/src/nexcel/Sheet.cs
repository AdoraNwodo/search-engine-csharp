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
using NExcel.Format;
namespace NExcel
{
	
	/// <summary> Represents a sheet within a workbook.  Provides a handle to the individual
	/// cells, or lines of cells (grouped by Row or Column)
	/// </summary>
	public interface Sheet
		{
			/// <summary> Returns the number of rows in this sheet
			/// 
			/// </summary>
			/// <returns> the number of rows in this sheet
			/// </returns>
			int Rows
			{
				get;
				
			}
			/// <summary> Returns the number of columns in this sheet
			/// 
			/// </summary>
			/// <returns> the number of columns in this sheet
			/// </returns>
			int Columns
			{
				get;
				
			}
			/// <summary> Gets the name of this sheet
			/// 
			/// </summary>
			/// <returns> the name of the sheet
			/// </returns>
			string Name
			{
				get;
				
			}
			/// <summary> Determines whether the sheet is hidden
			/// 
			/// </summary>
			/// <returns> whether or not the sheet is hidden
			/// </returns>
			/// <deprecated> in favouf of the getSettings() method
			/// </deprecated>
			bool Hidden
			{
				get;
				
			}
			/// <summary> Determines whether the sheet is protected
			/// 
			/// </summary>
			/// <returns> whether or not the sheet is protected
			/// </returns>
			/// <deprecated> in favour of the getSettings() method
			/// </deprecated>
			bool Protected
			{
				get;
				
			}
			/// <summary> Gets the hyperlinks on this sheet
			/// 
			/// </summary>
			/// <returns> an array of hyperlinks
			/// </returns>
			Hyperlink[] Hyperlinks
			{
				get;
				
			}
			/// <summary> Gets the cells which have been merged on this sheet
			/// 
			/// </summary>
			/// <returns> an array of range objects
			/// </returns>
			Range[] MergedCells
			{
				get;
				
			}
			/// <summary> Gets the settings used on a particular sheet
			/// 
			/// </summary>
			/// <returns> the sheet settings
			/// </returns>
			SheetSettings Settings
			{
				get;
				
			}
			/// <summary> Returns the cell specified at this row and at this column.
			/// If a column/row combination forms part of a merged group of cells
			/// then (unless it is the first cell of the group) a blank cell
			/// will be returned
			/// 
			/// </summary>
			/// <param name="column">the column number
			/// </param>
			/// <param name="row">the row number
			/// </param>
			/// <returns> the cell at the specified co-ordinates
			/// </returns>
			Cell getCell(int column, int row);
			
			/// <summary> Gets all the cells on the specified row
			/// 
			/// </summary>
			/// <param name="row">the rows whose cells are to be returned
			/// </param>
			/// <returns> the cells on the given row
			/// </returns>
			Cell[] getRow(int row);
			
			/// <summary> Gets all the cells on the specified column
			/// 
			/// </summary>
			/// <param name="col">the column whose cells are to be returned
			/// </param>
			/// <returns> the cells on the specified column
			/// </returns>
			Cell[] getColumn(int col);
			
			/// <summary> Gets the cell whose contents match the string passed in.
			/// If no match is found, then null is returned.  The search is performed
			/// on a row by row basis, so the lower the row number, the more
			/// efficiently the algorithm will perform
			/// 
			/// </summary>
			/// <param name="contents">the string to match
			/// </param>
			/// <returns> the Cell whose contents match the paramter, null if not found
			/// </returns>
			Cell findCell(string contents);
			
			/// <summary> Gets the cell whose contents match the string passed in.
			/// If no match is found, then null is returned.  The search is performed
			/// on a row by row basis, so the lower the row number, the more
			/// efficiently the algorithm will perform.  This method differs
			/// from the findCell method in that only cells with labels are
			/// queried - all numerical cells are ignored.  This should therefore
			/// improve performance.
			/// 
			/// </summary>
			/// <param name="contents">the string to match
			/// </param>
			/// <returns> the Cell whose contents match the paramter, null if not found
			/// </returns>
			LabelCell findLabelCell(string contents);
			
			/// <summary> Gets the column format for the specified column
			/// 
			/// </summary>
			/// <param name="col">the column number
			/// </param>
			/// <returns> the column format, or NULL if the column has no specific format
			/// </returns>
			/// <deprecated> Use getColumnView and the CellView bean instead
			/// </deprecated>
			NExcel.Format.CellFormat getColumnFormat(int col);
			
			/// <summary> Gets the column width for the specified column
			/// 
			/// </summary>
			/// <param name="col">the column number
			/// </param>
			/// <returns> the column width, or the default width if the column has no
			/// specified format
			/// </returns>
			/// <deprecated> Use getColumnView instead
			/// </deprecated>
			int getColumnWidth(int col);
			
			/// <summary> Gets the column width for the specified column
			/// 
			/// </summary>
			/// <param name="col">the column number
			/// </param>
			/// <returns> the column format, or the default format if no override is
			/// specified
			/// </returns>
			CellView getColumnView(int col);
			
			/// <summary> Gets the column width for the specified column
			/// 
			/// </summary>
			/// <param name="row">the column number
			/// </param>
			/// <returns> the row height, or the default height if the column has no
			/// specified format
			/// </returns>
			/// <deprecated> use getRowView instead
			/// </deprecated>
			int getRowHeight(int row);
			
			/// <summary> Gets the column width for the specified column
			/// 
			/// </summary>
			/// <param name="row">the column number
			/// </param>
			/// <returns> the row format, which may be the default format if no format
			/// is specified
			/// </returns>
			CellView getRowView(int row);
		}
}
