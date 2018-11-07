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
using common;
using NExcel;
using NExcel.Biff;
namespace NExcel.Biff.Formula
{
	
	/// <summary> A cell reference in a formula</summary>
	class IntegerValue:NumberValue, ParsedThing
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
				sbyte[] data = new sbyte[3];
				data[0] = Token.INTEGER.Code;
				
				IntegerHelper.getTwoBytes((int) Value, data, 1);
				
				return data;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The value of this integer</summary>
		private double Value;
		
		/// <summary> Constructor</summary>
		public IntegerValue()
		{
		}
		
		/// <summary> Constructor for an integer value being read from a string</summary>
		public IntegerValue(string s)
		{
			try
			{
				Value = System.Int32.Parse(s);
			}
			catch (System.FormatException e)
			{
				logger.warn(e, e);
				Value = 0;
			}
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
		public override int read(sbyte[] data, int pos)
		{
			Value = IntegerHelper.getInt(data[pos], data[pos + 1]);
			
			return 2;
		}
		
		/// <summary> Accessor for the value
		/// 
		/// </summary>
		/// <returns> the value
		/// </returns>
		public override double getValue()
		{
			return Value;
		}
		static IntegerValue()
		{
			logger = Logger.getLogger(typeof(IntegerValue));
		}
	}
}
