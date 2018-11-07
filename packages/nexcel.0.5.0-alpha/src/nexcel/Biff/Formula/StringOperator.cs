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
namespace NExcel.Biff.Formula
{
	
	/// <summary> Ambiguously defined operator, used as a place holder when parsing
	/// string formulas.  At this stage it could be either
	/// a unary or binary operator - the string parser will deduce which and
	/// create the appropriate type
	/// </summary>
	abstract class StringOperator:Operator
	{
		/// <summary> Gets the precedence for this operator.  Does nothing here
		/// 
		/// </summary>
		/// <returns> the operator precedence
		/// </returns>
		override internal int Precedence
		{
			get
			{
				Assert.verify(false);
				return 0;
			}
			
		}
		/// <summary> Gets the token representation of this item in RPN.  Does nothing here
		/// 
		/// </summary>
		/// <returns> the bytes applicable to this formula
		/// </returns>
		override internal sbyte[] Bytes
		{
			get
			{
				Assert.verify(false);
				return null;
			}
			
		}
		/// <summary> Abstract method which gets the binary version of this operator</summary>
		internal abstract Operator BinaryOperator{get;}
		/// <summary> Abstract method which gets the unary version of this operator</summary>
		internal abstract Operator UnaryOperator{get;}
		/// <summary> Constructor</summary>
		protected internal StringOperator():base()
		{
		}
		
		/// <summary> Gets the operands for this operator from the stack.  Does nothing
		/// here
		/// </summary>
		public override void getOperands(Stack s)
		{
		Assert.verify(false);
		}
		
		/// <summary> Gets the string representation of this item</summary>
		public override void  getString(System.Text.StringBuilder buf)
		{
			Assert.verify(false);
		}
		
		/// <summary> Default behaviour is to do nothing
		/// 
		/// </summary>
		/// <param name="colAdjust">the amount to add on to each relative cell reference
		/// </param>
		/// <param name="rowAdjust">the amount to add on to each relative row reference
		/// </param>
		public override void  adjustRelativeCellReferences(int colAdjust, int rowAdjust)
		{
			Assert.verify(false);
		}
		
		
		/// <summary> Default behaviour is to do nothing
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
			Assert.verify(false);
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
			Assert.verify(false);
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
			Assert.verify(false);
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
			Assert.verify(false);
		}
	}
}
