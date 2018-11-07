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
	
	/// <summary> The environment specific print record</summary>
	public class PLSRecord:RecordData
	{
		/// <summary> Gets the data
		/// 
		/// </summary>
		/// <returns> the binary data
		/// </returns>
		virtual public sbyte[] Data
		{
			get
			{
				return getRecord().Data;
			}
			
		}
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="r">the raw data
		/// </param>
		public PLSRecord(Record r):base(r)
		{
		}
	}
}
