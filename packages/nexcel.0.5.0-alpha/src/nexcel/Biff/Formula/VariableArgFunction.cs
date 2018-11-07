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
	
	/// <summary> A built in function in a formula.  These functions take a variable
	/// number of arguments, such as a range (eg. SUM etc)
	/// </summary>
	class VariableArgFunction:Operator, ParsedThing
	{
		/// <summary> Gets the underlying function</summary>
		virtual internal Function Function
		{
			get
			{
				return function;
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
				handleSpecialCases();
				
				// Get the data for the operands - in reverse order
				ParseItem[] operands = getOperands();
				sbyte[] data = new sbyte[0];
				
				for (int i = operands.Length - 1; i >= 0; i--)
				{
					sbyte[] opdata = operands[i].Bytes;
					
					// Grow the array
					sbyte[] cnewdata = new sbyte[data.Length + opdata.Length];
					Array.Copy(data, 0, cnewdata, 0, data.Length);
					Array.Copy(opdata, 0, cnewdata, data.Length, opdata.Length);
					data = cnewdata;
				}
				
				// Add on the operator byte
				sbyte[] newdata = new sbyte[data.Length + 4];
				Array.Copy(data, 0, newdata, 0, data.Length);
				newdata[data.Length] = !useAlternateCode()?Token.FUNCTIONVARARG.Code:Token.FUNCTIONVARARG.Code2;
				newdata[data.Length + 1] = (sbyte) arguments;
				IntegerHelper.getTwoBytes(function.Code, newdata, data.Length + 2);
				
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
		
		/// <summary> The number of arguments</summary>
		private int arguments;
		
		/// <summary> Flag which indicates whether this was initialized from the client
		/// api or from an excel sheet
		/// </summary>
		private bool readFromSheet;
		
		/// <summary> The workbooks settings</summary>
		private WorkbookSettings settings;
		
		/// <summary> Constructor</summary>
		public VariableArgFunction(WorkbookSettings ws)
		{
			readFromSheet = true;
			settings = ws;
		}
		
		/// <summary> Constructor used when parsing a function from a string
		/// 
		/// </summary>
		/// <param name="f">the function
		/// </param>
		/// <param name="a">the number of arguments
		/// </param>
		public VariableArgFunction(Function f, int a, WorkbookSettings ws)
		{
			function = f;
			arguments = a;
			readFromSheet = false;
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
		/// <exception cref=""> FormulaException
		/// </exception>
		public virtual int read(sbyte[] data, int pos)
		{
			arguments = data[pos];
			int index = IntegerHelper.getInt(data[pos + 1], data[pos + 2]);
			function = Function.getFunction(index);
			
			if (function == NExcel.Biff.Formula.Function.UNKNOWN)
			{
				throw new FormulaException(FormulaException.unrecognizedFunction, index);
			}
			
			return 3;
		}
		
		/// <summary> Gets the operands for this operator from the stack</summary>
		public override void  getOperands(Stack s)
		{
		// parameters are in the correct order, god damn them
		ParseItem[] items = new ParseItem[arguments];
		
		for (int i = arguments - 1; i >= 0; i--)
		{
		ParseItem pi = (ParseItem) s.Pop();
		
		items[i] = pi;
		}
		
		for (int i = 0; i < arguments; i++)
		{
		add(items[i]);
		}
		}
		
		
		public override void  getString(System.Text.StringBuilder buf)
		{
			buf.Append(function.getName(settings));
			buf.Append('(');
			
			if (arguments > 0)
			{
				ParseItem[] operands = getOperands();
				if (readFromSheet)
				{
					// arguments are in the same order they were specified
					operands[0].getString(buf);
					
					for (int i = 1; i < arguments; i++)
					{
						buf.Append(',');
						operands[i].getString(buf);
					}
				}
				else
				{
					// arguments are stored in the reverse order to which they
					// were specified, so iterate through them backwards
					operands[arguments - 1].getString(buf);
					
					for (int i = arguments - 2; i >= 0; i--)
					{
						buf.Append(',');
						operands[i].getString(buf);
					}
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
		
		/// <summary> Handles functions which form a special case</summary>
		private void  handleSpecialCases()
		{
			// Handle the array functions.  Tell all the Area records to
			// use their alternative token code
			if (function == Function.SUMPRODUCT)
			{
				// Get the data for the operands - in reverse order
				ParseItem[] operands = getOperands();
				
				for (int i = operands.Length - 1; i >= 0; i--)
				{
					if (operands[i] is Area)
					{
						operands[i].setAlternateCode();
					}
				}
			}
		}
		static VariableArgFunction()
		{
			logger = Logger.getLogger(typeof(VariableArgFunction));
		}
	}
}
