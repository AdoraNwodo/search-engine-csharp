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
	/// <summary> Indicates that the function doesn't evaluate to a constant reference</summary>
	
	class MemFunc:Operand, ParsedThing
	{
		/// <summary> Gets the token representation of this item in RPN.  The Attribute
		/// token is a special case, which overrides anything useful we could do
		/// in the base class
		/// 
		/// </summary>
		/// <returns> the bytes applicable to this formula
		/// </returns>
		override internal sbyte[] Bytes
		{
			get
			{
				return null;
			}
			
		}
		/// <summary> Gets the precedence for this operator.  Operator precedents run from 
		/// 1 to 5, one being the highest, 5 being the lowest
		/// 
		/// </summary>
		/// <returns> the operator precedence
		/// </returns>
		virtual internal int Precedence
		{
			get
			{
				return 5;
			}
			
		}
		/// <summary> Accessor for the .Length
		/// 
		/// </summary>
		/// <returns> the .Length of the subexpression
		/// </returns>
		virtual public int Length
		{
			get
			{
				return length;
			}
			
		}
		virtual public ParseItem[] SubExpression
		{
			set
			{
				subExpression = value;
			}
			
		}
		/// <summary> The number of bytes in the subexpression</summary>
		private int length;
		
		/// <summary> The sub expression</summary>
		private ParseItem[] subExpression;
		
		/// <summary> Constructor</summary>
		public MemFunc()
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
			length = IntegerHelper.getInt(data[pos], data[pos + 1]);
			return 2;
		}
		
		/// <summary> Gets the operands for this operator from the stack</summary>
		public virtual void  getOperands(Stack s)
		{
		}
		
		public override void  getString(System.Text.StringBuilder buf)
		{
			if (subExpression.Length == 1)
			{
				subExpression[0].getString(buf);
			}
			else if (subExpression.Length == 2)
			{
				subExpression[1].getString(buf);
				buf.Append(':');
				subExpression[0].getString(buf);
			}
		}
	}
}
