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

namespace NExcel.Read.Biff
{
	
	/// <summary> Helper to convert an RK number into a double or an integer</summary>
	sealed class RKHelper
	{
		/// <summary> Private constructor to prevent instantiation</summary>
		private RKHelper()
		{
		}
		
		/// <summary> Converts excel's internal RK format into a double value
		/// 
		/// </summary>
		/// <param name="rk">the rk number in bits
		/// </param>
		/// <returns> the double representation
		/// </returns>
		public static double getDouble(int rk)
		{
			if ((rk & 0x02) != 0)
			{
				int intval = rk >> 2;
				
				double Value = (double) intval;
				if ((rk & 0x01) != 0)
				{
					Value /= 100;
				}
				
				return Value;
			}
			else
			{
				long valbits = ( ((long) rk) & 0xfffffffc);
				valbits <<= 32;
				
				double Value = BitConverter.Int64BitsToDouble(valbits);
				
				
				if ((rk & 0x01) != 0)
				{
					Value /= 100;
				}
				
				return Value;
			}
		}
	}
}
