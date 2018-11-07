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
	/// <summary> A name operand</summary>
	class Name:Operand, ParsedThing
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
				sbyte[] data = new sbyte[6];
				
				return data;
			}
			
		}
		/// <summary> Constructor</summary>
		public Name()
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
			int ixti = IntegerHelper.getInt(data[pos], data[pos + 1]);
			int ilbl = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
			
			return 6;
		}
		
		/// <summary> Abstract method implementation to get the string equivalent of this
		/// token
		/// 
		/// </summary>
		/// <param name="buf">the string to append to
		/// </param>
		public override void  getString(System.Text.StringBuilder buf)
		{
			buf.Append("[Name record not implemented]");
		}
	}
}
