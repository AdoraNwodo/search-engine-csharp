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
//import NExcel.Workbook;
using NExcel.Biff;
namespace NExcel.Biff.Formula
{
	
	/// <summary> Parses the formula passed in (either as parsed strings or as a string)
	/// into a tree of operators and operands
	/// </summary>
	public class FormulaParser
	{
		/// <summary> Gets the formula as a string
		/// 
		/// </summary>
		/// <exception cref=""> FormulaException
		/// </exception>
		virtual public string Formula
		{
			get
			{
				return parser.Formula;
			}
			
		}
		/// <summary> Gets the bytes for the formula. This takes into account any
		/// token mapping necessary because of shared formulas
		/// 
		/// </summary>
		/// <returns> the bytes in RPN
		/// </returns>
		virtual public sbyte[] Bytes
		{
			get
			{
				return parser.Bytes;
			}
			
		}
		/// <summary> The formula parser.  The object implementing this interface will either
		/// parse tokens or strings
		/// </summary>
		private Parser parser;
		
		/// <summary> Constructor which creates the parse tree out of tokens
		/// 
		/// </summary>
		/// <param name="tokens">the list of parsed tokens
		/// </param>
		/// <param name="rt">the cell containing the formula
		/// </param>
		/// <param name="es">a handle to the external sheet
		/// </param>
		/// <param name="nt">a handle to the name table
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <exception cref=""> FormulaException
		/// </exception>
		public FormulaParser(sbyte[] tokens, Cell rt, ExternalSheet es, WorkbookMethods nt, WorkbookSettings ws)
		{
			// A null workbook bof means that it is a writable workbook and therefore
			// must be biff8
			if (es.WorkbookBof != null && !es.WorkbookBof.isBiff8())
			{
				throw new FormulaException(FormulaException.biff8Supported);
			}
			parser = new TokenFormulaParser(tokens, rt, es, nt, ws);
		}
		
		/// <summary> Constructor which creates the parse tree out of the string
		/// 
		/// </summary>
		/// <param name="form">the formula string
		/// </param>
		/// <param name="es">the external sheet handle
		/// </param>
		/// <param name="nt">the name table
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public FormulaParser(string form, ExternalSheet es, WorkbookMethods nt, WorkbookSettings ws)
		{
			parser = new StringFormulaParser(form, es, nt, ws);
		}
		
		
		/// <summary> Adjusts all the relative cell references in this formula by the
		/// amount specified.  Used when copying formulas
		/// 
		/// </summary>
		/// <param name="colAdjust">the amount to add on to each relative cell reference
		/// </param>
		/// <param name="rowAdjust">the amount to add on to each relative row reference
		/// </param>
		public virtual void  adjustRelativeCellReferences(int colAdjust, int rowAdjust)
		{
			parser.adjustRelativeCellReferences(colAdjust, rowAdjust);
		}
		
		/// <summary> Parses the formula into a parse tree
		/// 
		/// </summary>
		/// <exception cref=""> FormulaException
		/// </exception>
		public virtual void  parse()
		{
			parser.parse();
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
		public virtual void  columnInserted(int sheetIndex, int col, bool currentSheet)
		{
			parser.columnInserted(sheetIndex, col, currentSheet);
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
		public virtual void  columnRemoved(int sheetIndex, int col, bool currentSheet)
		{
			parser.columnRemoved(sheetIndex, col, currentSheet);
		}
		
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
		public virtual void  rowInserted(int sheetIndex, int row, bool currentSheet)
		{
			parser.rowInserted(sheetIndex, row, currentSheet);
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
		public virtual void  rowRemoved(int sheetIndex, int row, bool currentSheet)
		{
			parser.rowRemoved(sheetIndex, row, currentSheet);
		}
	}
}
