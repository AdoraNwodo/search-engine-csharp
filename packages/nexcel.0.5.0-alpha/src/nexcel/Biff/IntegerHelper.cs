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
	
	/// <summary> Converts excel byte representations into integers</summary>
	public sealed class IntegerHelper
	{
		/// <summary> Private constructor disables the instantiation of this object</summary>
		private IntegerHelper()
		{
		}
		
		/// <summary> Gets an int from two bytes
		/// 
		/// </summary>
		/// <param name="b2">the second byte
		/// </param>
		/// <param name="b1">the first byte
		/// </param>
		/// <returns> The integer value
		/// </returns>
		public static int getInt(sbyte b1, sbyte b2)
		{
			int i1 = ((int) b1) & 0xff;
			int i2 = ((int) b2) & 0xff;
			int val = i2 << 8 | i1;
			return val;
		}
		
		/// <summary> Gets an short from two bytes
		/// 
		/// </summary>
		/// <param name="b2">the second byte
		/// </param>
		/// <param name="b1">the first byte
		/// </param>
		/// <returns> The short value
		/// </returns>
		public static short getShort(sbyte b1, sbyte b2)
		{
			short i1 = (short) (b1 & 0xff);
			short i2 = (short) (b2 & 0xff);
			short val = (short) (i2 << 8 | i1);
			return val;
		}
		
		
		/// <summary> Gets an int from four bytes, doing all the necessary swapping
		/// 
		/// </summary>
		/// <param name="b1">a byte
		/// </param>
		/// <param name="b2">a byte
		/// </param>
		/// <param name="b3">a byte
		/// </param>
		/// <param name="b4">a byte
		/// </param>
		/// <returns> the integer value represented by the four bytes
		/// </returns>
		public static int getInt(sbyte b1, sbyte b2, sbyte b3, sbyte b4)
		{
			int i1 = getInt(b1, b2);
			int i2 = getInt(b3, b4);
			
			int val = i2 << 16 | i1;
			return val;
		}
		
		/// <summary> Gets a two byte array from an integer
		/// 
		/// </summary>
		/// <param name="i">the integer
		/// </param>
		/// <returns> the two bytes
		/// </returns>
		public static sbyte[] getTwoBytes(int i)
		{
			sbyte[] bytes = new sbyte[2];
			
			bytes[0] = (sbyte) (i & 0xff);
			bytes[1] = (sbyte) ((i & 0xff00) >> 8);
			
			return bytes;
		}
		
		/// <summary> Gets a four byte array from an integer
		/// 
		/// </summary>
		/// <param name="i">the integer
		/// </param>
		/// <returns> a four byte array
		/// </returns>
		public static sbyte[] getFourBytes(int i)
		{
			sbyte[] bytes = new sbyte[4];
			
			int i1 = i & 0xffff;
			int i2 = (int) ( (((long) i) & 0xffff0000) >> 16);
			
			getTwoBytes(i1, bytes, 0);
			getTwoBytes(i2, bytes, 2);
			
			return bytes;
		}
		
		
		/// <summary> Converts an integer into two bytes, and places it in the array at the
		/// specified position
		/// 
		/// </summary>
		/// <param name="target">the array to place the byte data into
		/// </param>
		/// <param name="pos">the position at which to place the data
		/// </param>
		/// <param name="i">the integer value to convert
		/// </param>
		public static void  getTwoBytes(int i, sbyte[] target, int pos)
		{
			sbyte[] bytes = getTwoBytes(i);
			target[pos] = bytes[0];
			target[pos + 1] = bytes[1];
		}
		
		/// <summary> Converts an integer into four bytes, and places it in the array at the
		/// specified position
		/// 
		/// </summary>
		/// <param name="target">the array which is to contain the converted data
		/// </param>
		/// <param name="pos">the position in the array in which to place the data
		/// </param>
		/// <param name="i">the integer to convert
		/// </param>
		public static void  getFourBytes(int i, sbyte[] target, int pos)
		{
			sbyte[] bytes = getFourBytes(i);
			target[pos] = bytes[0];
			target[pos + 1] = bytes[1];
			target[pos + 2] = bytes[2];
			target[pos + 3] = bytes[3];
		}
	}
}
