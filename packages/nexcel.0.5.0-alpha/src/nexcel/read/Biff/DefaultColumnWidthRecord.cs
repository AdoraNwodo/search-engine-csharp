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
namespace NExcel.Read.Biff
{
	
	/// <summary> Contains the default column width for cells in this sheet</summary>
	class DefaultColumnWidthRecord:RecordData
	{
		/// <summary> Accessor for the default width
		/// 
		/// </summary>
		/// <returns> the width
		/// </returns>
		virtual public int Width
		{
			get
			{
				return width;
			}
			
		}
		/// <summary> The default columns width, in characters</summary>
		private int width;
		
		/// <summary> Constructs the def col width from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public DefaultColumnWidthRecord(Record t):base(t)
		{
			sbyte[] data = t.Data;
			
			width = IntegerHelper.getInt(data[0], data[1]);
		}
	}
}
