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
using NExcel.Biff;
namespace NExcel.Biff.Formula
{
	
	/// <summary> A cell reference in a formula</summary>
	class Power:BinaryOperator, ParsedThing
	{
		/// <summary> Abstract method which gets the token for this operator
		/// 
		/// </summary>
		/// <returns> the string symbol for this token
		/// </returns>
		override internal Token Token
		{
			get
			{
				return Token.POWER;
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
				return 1;
			}
			
		}
		/// <summary> Constructor</summary>
		public Power()
		{
		}
		
		public override string getSymbol()
		{
			return "^";
		}
	}
}
