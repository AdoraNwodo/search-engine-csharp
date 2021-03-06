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
using NExcel.Biff;
namespace NExcel.Biff.Formula
{
	
	/// <summary> A cell reference in a formula</summary>
	abstract class BinaryOperator:Operator, ParsedThing
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
				
				// Get the operands in reverse order to get the RPN
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
				sbyte[] newdata = new sbyte[data.Length + 1];
				Array.Copy(data, 0, newdata, 0, data.Length);
				newdata[data.Length] = Token.Code;
				
				return newdata;
			}
			
		}
		/// <summary> Abstract method which gets the token for this operator
		/// 
		/// </summary>
		/// <returns> the string symbol for this token
		/// </returns>
		internal abstract Token Token{get;}
		/// <summary> Constructor</summary>
		public BinaryOperator()
		{
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
			return 0;
		}
		
		/// <summary> Gets the operands for this operator from the stack</summary>
		public override void  getOperands(Stack s)
		{
			ParseItem o1 = (ParseItem) s.Pop();
			ParseItem o2 = (ParseItem) s.Pop();
			
			add(o1);
			add(o2);
		}
		
		/// <summary> Gets the string version of this binary operator
		/// 
		/// </summary>
		/// <param name="buf">a the string buffer
		/// </param>
		public override void  getString(System.Text.StringBuilder buf)
		{
			ParseItem[] operands = getOperands();
			operands[1].getString(buf);
			buf.Append(getSymbol());
			operands[0].getString(buf);
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
			operands[1].adjustRelativeCellReferences(colAdjust, rowAdjust);
			operands[0].adjustRelativeCellReferences(colAdjust, rowAdjust);
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
			operands[1].columnInserted(sheetIndex, col, currentSheet);
			operands[0].columnInserted(sheetIndex, col, currentSheet);
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
			operands[1].columnRemoved(sheetIndex, col, currentSheet);
			operands[0].columnRemoved(sheetIndex, col, currentSheet);
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
			operands[1].rowInserted(sheetIndex, row, currentSheet);
			operands[0].rowInserted(sheetIndex, row, currentSheet);
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
			operands[1].rowRemoved(sheetIndex, row, currentSheet);
			operands[0].rowRemoved(sheetIndex, row, currentSheet);
		}
		
		/// <summary> Abstract method which gets the binary operator string symbol
		/// 
		/// </summary>
		/// <returns> the string symbol for this token
		/// </returns>
		public abstract string getSymbol();
	}
}
