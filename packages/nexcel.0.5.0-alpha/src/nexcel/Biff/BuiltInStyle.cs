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
	
	/// <summary> Represents a built in, rather than a user defined, style.
	/// This class is used by the FormattingRecords class when writing out the hard*
	/// coded styles
	/// </summary>
	class BuiltInStyle:WritableRecordData
	{
		/// <summary> The XF index of this style</summary>
		private int xfIndex;
		/// <summary> The reference number of this style</summary>
		private int styleNumber;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="xfind">the xf index of this style
		/// </param>
		/// <param name="sn">the style number of this style
		/// </param>
		public BuiltInStyle(int xfind, int sn):base(NExcel.Biff.Type.STYLE)
		{
			
			xfIndex = xfind;
			styleNumber = sn;
		}
		
		/// <summary> Abstract method implementation to get the raw byte data ready to write out
		/// 
		/// </summary>
		/// <returns> The byte data
		/// </returns>
		public override sbyte[] getData()
		{
			sbyte[] data = new sbyte[4];
			
			IntegerHelper.getTwoBytes(xfIndex, data, 0);
			
			// Set the built in bit
			data[1] |= (sbyte) -0x80;
			
			data[2] = (sbyte) styleNumber;
			
			// Set the outline level
			data[3] = (sbyte) -0x01;
			
			return data;
		}
	}
}
