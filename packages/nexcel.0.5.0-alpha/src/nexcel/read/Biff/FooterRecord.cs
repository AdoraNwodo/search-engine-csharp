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
using NExcel;
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	/// <summary> A workbook page footer record</summary>
	public class FooterRecord:RecordData
	{
		/// <summary> Gets the footer string
		/// 
		/// </summary>
		/// <returns> the footer string
		/// </returns>
		virtual internal string Footer
		{
			get
			{
				return footer;
			}
			
		}
		/// <summary> The footer</summary>
		private string footer;
		
		/// <summary> Dummy indicators for overloading the constructor</summary>
		public class Biff7
		{
		}
		
		public static Biff7 biff7;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the record data
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		internal FooterRecord(Record t, WorkbookSettings ws):base(t)
		{
			sbyte[] data = getRecord().Data;
			
			if (data.Length == 0)
			{
				return ;
			}
			
			int chars = IntegerHelper.getInt(data[0], data[1]);
			
			bool unicode = data[2] == 1;
			
			if (unicode)
			{
				footer = StringHelper.getUnicodeString(data, chars, 3);
			}
			else
			{
				footer = StringHelper.getString(data, chars, 3, ws);
			}
		}
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the record data
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <param name="dummy">dummy record to indicate a biff7 document
		/// </param>
		internal FooterRecord(Record t, WorkbookSettings ws, Biff7 dummy):base(t)
		{
			sbyte[] data = getRecord().Data;
			
			if (data.Length == 0)
			{
				return ;
			}
			
			int chars = data[0];
			footer = StringHelper.getString(data, chars, 1, ws);
		}
		static FooterRecord()
		{
			biff7 = new Biff7();
		}
	}
}
