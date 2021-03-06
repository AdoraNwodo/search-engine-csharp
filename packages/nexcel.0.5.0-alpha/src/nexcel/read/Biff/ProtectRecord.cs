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
	
	/// <summary> A record detailing whether the sheet is protected</summary>
	class ProtectRecord:RecordData
	{
		/// <summary> Returns the protected flag
		/// 
		/// </summary>
		/// <returns> TRUE if this is protected, FALSE otherwise
		/// </returns>
		virtual internal bool IsProtected()
		{
				return prot;
		}
		/// <summary> Protected flag</summary>
		private bool prot;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		internal ProtectRecord(Record t):base(t)
		{
			sbyte[] data = getRecord().Data;
			
			int protflag = IntegerHelper.getInt(data[0], data[1]);
			
			prot = (protflag == 1);
		}
	}
}
