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
namespace NExcel.Biff.Formula
{
	
	/// <summary> Interface used by the two different types of formula parser</summary>
	internal interface Parser
		{
			/// <summary> Gets the string version of the formula
			/// 
			/// </summary>
			/// <returns> the formula as a string
			/// </returns>
			string Formula
			{
				get;
				
			}
			/// <summary> Gets the bytes for the formula. This takes into account any
			/// token mapping necessary because of shared formulas
			/// 
			/// </summary>
			/// <returns> the bytes in RPN
			/// </returns>
			sbyte[] Bytes
			{
				get;
				
			}
			/// <summary> Parses the formula
			/// 
			/// </summary>
			/// <exception cref=""> FormulaException if an error occurs
			/// </exception>
			void  parse();
			
			/// <summary> Adjusts all the relative cell references in this formula by the
			/// amount specified.  
			/// 
			/// </summary>
			/// <param name="">colAdjust
			/// </param>
			/// <param name="">rowAdjust
			/// </param>
			void  adjustRelativeCellReferences(int colAdjust, int rowAdjust);
			
			
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
			void  columnInserted(int sheetIndex, int col, bool currentSheet);
			
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
			void  columnRemoved(int sheetIndex, int col, bool currentSheet);
			
			/// <summary> Called when a column is inserted on the specified sheet.  Tells
			/// the formula  parser to update all of its cell references beyond this
			/// column
			/// 
			/// </summary>
			/// <param name="sheetIndex">the sheet on which the column was inserted
			/// </param>
			/// <param name="row">the column number which was inserted
			/// </param>
			/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
			/// column was inserted, FALSE otherwise
			/// </param>
			void  rowInserted(int sheetIndex, int row, bool currentSheet);
			
			/// <summary> Called when a column is inserted on the specified sheet.  Tells
			/// the formula  parser to update all of its cell references beyond this
			/// column
			/// 
			/// </summary>
			/// <param name="sheetIndex">the sheet on which the column was removed
			/// </param>
			/// <param name="row">the column number which was removed
			/// </param>
			/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
			/// column was inserted, FALSE otherwise
			/// </param>
			void  rowRemoved(int sheetIndex, int row, bool currentSheet);
		}
}
