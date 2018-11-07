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
using NExcel.Biff;
namespace NExcel.Biff.Formula
{
	
	/// <summary> A boolean operand in a formula</summary>
	class BooleanValue:Operand, ParsedThing
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
				sbyte[] data = new sbyte[2];
				data[0] = Token.BOOL.Code;
				data[1] = (sbyte) (Value == true?1:0);
				
				return data;
			}
			
		}
		/// <summary> The boolean value</summary>
		private bool Value;
		
		/// <summary> Constructor</summary>
		public BooleanValue()
		{
		}
		
		/// <summary> Constructor used when parsing a string formula
		/// 
		/// </summary>
		/// <param name="s">the string token, including quote marks
		/// </param>
		public BooleanValue(string s)
		{
			// remove the quotes
			Value = System.Boolean.Parse(s);
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
			Value = data[pos] == 1?true:false;
			return 1;
		}
		
		/// <summary> Abstract method implementation to get the string equivalent of this
		/// token
		/// 
		/// </summary>
		/// <param name="buf">the string to append to
		/// </param>
		public override void  getString(System.Text.StringBuilder buf)
		{
			buf.Append(Value.ToString());
			
		}
	}
}
