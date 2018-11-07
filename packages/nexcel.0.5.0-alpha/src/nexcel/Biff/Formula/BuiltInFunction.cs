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
using NExcel;
using NExcel.Biff;
namespace NExcel.Biff.Formula
{
	
	/// <summary> A built in function in a formula</summary>
	class BuiltInFunction:Operator, ParsedThing
	{
		/// <summary> Gets the token representation of this item in RPN
		/// 
		/// </summary>
		/// <returns> the bytes applicable to this formula
		/// </returns>
		override internal sbyte[] Bytes
		{
			get
			{
				// Get the data for the operands
				ParseItem[] operands = getOperands();
				sbyte[] data = new sbyte[0];
				
				for (int i = 0; i < operands.Length; i++)
				{
					sbyte[] opdata = operands[i].Bytes;
					
					// Grow the array
					sbyte[] cnewdata = new sbyte[data.Length + opdata.Length];
					Array.Copy(data, 0, cnewdata, 0, data.Length);
					Array.Copy(opdata, 0, cnewdata, data.Length, opdata.Length);
					data = cnewdata;
				}
				
				// Add on the operator byte
				sbyte[] newdata = new sbyte[data.Length + 3];
				Array.Copy(data, 0, newdata, 0, data.Length);
				newdata[data.Length] = !useAlternateCode()?Token.FUNCTION.Code:Token.FUNCTION.Code2;
				IntegerHelper.getTwoBytes(function.Code, newdata, data.Length + 1);
				
				return newdata;
			}
			
		}
		/// <summary> Gets the precedence for this operator.  Operator precedents run from 
		/// 1 to 5, one being the highest, 5 being the lowest
		/// 
		/// </summary>
		/// <returns> the operator precedence
		/// </returns>
		override internal int Precedence
		{
			get
			{
				return 3;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The function</summary>
		private Function function;
		
		/// <summary> The workbook settings</summary>
		private WorkbookSettings settings;
		
		/// <summary> Constructor</summary>
		/// <param name="ws">the workbook settings
		/// </param>
		public BuiltInFunction(WorkbookSettings ws)
		{
			settings = ws;
		}
		
		/// <summary> Constructor used when parsing a formula from a string
		/// 
		/// </summary>
		/// <param name="f">the function
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public BuiltInFunction(Function f, WorkbookSettings ws)
		{
			function = f;
			settings = ws;
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
			int index = IntegerHelper.getInt(data[pos], data[pos + 1]);
			function = Function.getFunction(index);
			Assert.verify(function != NExcel.Biff.Formula.Function.UNKNOWN, "function code " + index);
			return 2;
		}
		
		/// <summary> Gets the operands for this operator from the stack</summary>
		public override void getOperands(Stack s)
		{
		for (int i = 0; i < function.NumArgs; i++)
		{
		ParseItem pi = (ParseItem) s.Pop();
		add(pi);
		}
		}
		
		
		/// <summary> Gets the string for this functions
		/// 
		/// </summary>
		/// <param name="buf">the buffer to append
		/// </param>
		public override void  getString(System.Text.StringBuilder buf)
		{
			buf.Append(function.getName(settings));
			buf.Append('(');
			
			int numArgs = function.NumArgs;
			
			if (numArgs > 0)
			{
				ParseItem[] operands = getOperands();
				
				// operands are in the correct order 
				operands[numArgs - 1].getString(buf);
				
				for (int i = numArgs - 2; i >= 0; i--)
				{
					buf.Append(',');
					operands[i].getString(buf);
				}
			}
			
			buf.Append(')');
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
			ParseItem[] operands = getOperands();
			
			for (int i = 0; i < operands.Length; i++)
			{
				operands[i].adjustRelativeCellReferences(colAdjust, rowAdjust);
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
			ParseItem[] operands = getOperands();
			for (int i = 0; i < operands.Length; i++)
			{
				operands[i].columnInserted(sheetIndex, col, currentSheet);
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
			ParseItem[] operands = getOperands();
			for (int i = 0; i < operands.Length; i++)
			{
				operands[i].columnRemoved(sheetIndex, col, currentSheet);
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
			ParseItem[] operands = getOperands();
			for (int i = 0; i < operands.Length; i++)
			{
				operands[i].rowInserted(sheetIndex, row, currentSheet);
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
			ParseItem[] operands = getOperands();
			for (int i = 0; i < operands.Length; i++)
			{
				operands[i].rowRemoved(sheetIndex, row, currentSheet);
			}
		}
		static BuiltInFunction()
		{
			logger = Logger.getLogger(typeof(BuiltInFunction));
		}
	}
}
