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
namespace NExcel.Format
{
	
	/// <summary> Enumeration type which contains the available excel paper sizes and
	/// their codes
	/// </summary>
	public sealed class PaperSize
	{
		/// <summary> Accessor for the internal binary value association with this paper
		/// size
		/// 
		/// </summary>
		/// <returns> the internal value
		/// </returns>
		public int Value
		{
			get
			{
				return val;
			}
			
		}
		/// <summary> The excel encoding</summary>
		private int val;
		
		/// <summary> The paper sizes</summary>
		private static PaperSize[] paperSizes;
		
		/// <summary> Constructor</summary>
		private PaperSize(int v)
		{
			val = v;
			
			// Grow the array and add this to it
			PaperSize[] newarray = new PaperSize[paperSizes.Length + 1];
			Array.Copy((System.Array) paperSizes, 0, (System.Array) newarray, 0, paperSizes.Length);
			newarray[paperSizes.Length] = this;
			paperSizes = newarray;
		}
		
		private class Dummy
		{
		}
		
		private static readonly Dummy unknown = new Dummy();
		
		/// <summary> Constructor with a dummy parameter for unknown paper sizes</summary>
		private PaperSize(int v, Dummy u)
		{
			val = v;
		}
		
		/// <summary> Gets the paper size for a specific value
		/// 
		/// </summary>
		/// <param name="val">the value
		/// </param>
		/// <returns> the paper size
		/// </returns>
		public static PaperSize getPaperSize(int val)
		{
			bool found = false;
			int pos = 0;
			
			while (!found && pos < paperSizes.Length)
			{
				if (paperSizes[pos].Value == val)
				{
					found = true;
				}
				else
				{
					pos++;
				}
			}
			
			if (found)
			{
				return paperSizes[pos];
			}
			
			return new PaperSize(val, unknown);
		}
		
		/// <summary> A4</summary>
		public static PaperSize A4;
		
		/// <summary> Small A4</summary>
		public static PaperSize A4_SMALL;
		
		/// <summary> A5</summary>
		public static PaperSize A5;
		
		/// <summary> US Letter</summary>
		public static PaperSize LETTER;
		
		/// <summary> US Legal</summary>
		public static PaperSize LEGAL;
		
		/// <summary> A3</summary>
		public static PaperSize A3;
		static PaperSize()
		{
			paperSizes = new PaperSize[0];
			A4 = new PaperSize(0x9);
			A4_SMALL = new PaperSize(0xa);
			A5 = new PaperSize(0xb);
			LETTER = new PaperSize(0x1);
			LEGAL = new PaperSize(0x5);
			A3 = new PaperSize(0x8);
		}
	}
}