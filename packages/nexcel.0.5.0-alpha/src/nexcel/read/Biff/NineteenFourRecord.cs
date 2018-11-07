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
	
	/// <summary> Identifies the date system as the 1904 system or not</summary>
	class NineteenFourRecord:RecordData
	{
		/// <summary> The base year for dates</summary>
		private bool nineteenFour;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public NineteenFourRecord(Record t):base(t)
		{
			
			sbyte[] data = getRecord().Data;
			
			nineteenFour = data[0] == 1?true:false;
		}
		
		/// <summary> Accessor to see whether this spreadsheets dates are based around
		/// 1904
		/// 
		/// </summary>
		/// <returns> true if this workbooks dates are based around the 1904
		/// date system
		/// </returns>
		public virtual bool is1904()
		{
			return nineteenFour;
		}
	}
}
