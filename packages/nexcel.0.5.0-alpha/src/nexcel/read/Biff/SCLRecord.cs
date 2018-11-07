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
	
	/// <summary> Class containing the zoom factor for display</summary>
	class SCLRecord:RecordData
	{
		/// <summary> Accessor for the zoom factor
		/// 
		/// </summary>
		/// <returns> the zoom factor as the nearest integer percentage
		/// </returns>
		virtual public int ZoomFactor
		{
			get
			{
				return numerator * 100 / denominator;
			}
			
		}
		/// <summary> The numerator of the zoom</summary>
		private int numerator;
		
		/// <summary> The denominator of the zoom</summary>
		private int denominator;
		
		/// <summary> Constructs this record from the raw data</summary>
		/// <param name="r">the record
		/// </param>
		protected internal SCLRecord(Record r):base(NExcel.Biff.Type.SCL)
		{
			
			sbyte[] data = r.Data;
			
			numerator = IntegerHelper.getInt(data[0], data[1]);
			denominator = IntegerHelper.getInt(data[2], data[3]);
		}
	}
}
