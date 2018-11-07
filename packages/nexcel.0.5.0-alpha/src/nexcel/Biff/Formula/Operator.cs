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
namespace NExcel.Biff.Formula
{
	
	/// <summary> An operator is a node in a parse tree.  Its children can be other
	/// operators or operands
	/// Arithmetic operators and functions are all considered operators
	/// </summary>
	abstract class Operator:ParseItem
	{
		/// <summary> Gets the precedence for this operator.  Operator precedents run from 
		/// 1 to 5, one being the highest, 5 being the lowest
		/// 
		/// </summary>
		/// <returns> the operator precedence
		/// </returns>
		internal abstract int Precedence{get;}
		/// <summary> The items which this operator manipulates. There will be at most two</summary>
		private ParseItem[] operands;
		
		/// <summary> Constructor</summary>
		public Operator()
		{
			operands = new ParseItem[0];
		}
		
		/// <summary> Tells the operands to use the alternate code</summary>
		protected internal virtual void  setOperandAlternateCode()
		{
			for (int i = 0; i < operands.Length; i++)
			{
				operands[i].setAlternateCode();
			}
		}
		
		/// <summary> Adds operands to this item</summary>
		protected internal virtual void  add(ParseItem n)
		{
			n.Parent = this;
			
			// Grow the array
			ParseItem[] newOperands = new ParseItem[operands.Length + 1];
			Array.Copy(operands, 0, newOperands, 0, operands.Length);
			newOperands[operands.Length] = n;
			operands = newOperands;
		}
		
		/// <summary> Gets the operands for this operator from the stack </summary>
		public abstract void  getOperands(Stack s);
		
		/// <summary> Gets the operands ie. the children of the node</summary>
		protected internal virtual ParseItem[] getOperands()
		{
			return operands;
		}
	}
}
