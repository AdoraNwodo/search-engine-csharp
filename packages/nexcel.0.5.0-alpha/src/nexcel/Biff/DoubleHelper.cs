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

namespace NExcel.Biff
{
	
	/// <summary> Class to help handle doubles</summary>
	public class DoubleHelper
	{
		/// <summary> Private constructor to prevent instantiation</summary>
		private DoubleHelper()
		{
		}
		
		/// <summary> Gets the IEEE value from the byte array passed in
		/// 
		/// </summary>
		/// <param name="pos">the position in the data block which contains the double value
		/// </param>
		/// <param name="data">the data block containing the raw bytes
		/// </param>
		/// <returns> the double value converted from the raw data
		/// </returns>
		public static double getIEEEDouble(sbyte[] data, int pos)
		{
			int num1 = IntegerHelper.getInt(data[pos], data[pos + 1], data[pos + 2], data[pos + 3]);
			int num2 = IntegerHelper.getInt(data[pos + 4], data[pos + 5], data[pos + 6], data[pos + 7]);
			
			// Long.parseLong doesn't like the sign bit, so have to extract this
			// information and put it in at the end.  (Acknowledgment:  thanks
			// to Ruben for pointing this out)
			bool negative = ( ( ((long) num2) & 0x80000000) != 0 );
			
			// Thanks to Lyle for the following improved IEEE double processing
			long val = ((num2 & 0x7fffffff) * 0x100000000L) + (num1 < 0?0x100000000L + num1:num1);
			double Value = BitConverter.Int64BitsToDouble(val);
			
			if (negative)
			{
				Value = - Value;
			}
			return Value;
		}
		
		/// <summary> Puts the IEEE representation of the double provided into the array
		/// at the designated position
		/// 
		/// </summary>
		/// <param name="target">the data block into which the binary representation is to
		/// be placed
		/// </param>
		/// <param name="pos">the position in target in which to place the bytes
		/// </param>
		/// <param name="d">the double Value to convert to raw bytes
		/// </param>
		public static void  getIEEEBytes(double d, sbyte[] target, int pos)
		{
			long val = BitConverter.DoubleToInt64Bits(d);
			
			target[pos] = (sbyte) (val & 0xff);
			target[pos + 1] = (sbyte) ((val & 0xff00) >> 8);
			target[pos + 2] = (sbyte) ((val & 0xff0000) >> 16);
			target[pos + 3] = (sbyte) ((val & 0xff000000) >> 24);
			target[pos + 4] = (sbyte) ((val & 0xff00000000L) >> 32);
			target[pos + 5] = (sbyte) ((val & 0xff0000000000L) >> 40);
			target[pos + 6] = (sbyte) ((val & 0xff000000000000L) >> 48);
//			target[pos + 7] = (sbyte) ((val & 0xff00000000000000L) >> 56);
			target[pos + 7] = (sbyte) (val >> 56);
		}
	}
}
