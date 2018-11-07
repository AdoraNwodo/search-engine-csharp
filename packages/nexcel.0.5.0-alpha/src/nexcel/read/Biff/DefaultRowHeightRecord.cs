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
	class DefaultRowHeightRecord:RecordData
	{
		/// <summary> Accessor for the default height
		/// 
		/// </summary>
		/// <returns> the height
		/// </returns>
		virtual public int Height
		{
			get
			{
				return height;
			}
			
		}
		/// <summary> The default row height, in 1/20ths of a point</summary>
		private int height;
		
		/// <summary> Constructs the def col width from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public DefaultRowHeightRecord(Record t):base(t)
		{
			sbyte[] data = t.Data;
			
			if (data.Length > 2)
			{
				height = IntegerHelper.getInt(data[2], data[3]);
			}
		}
	}
}
