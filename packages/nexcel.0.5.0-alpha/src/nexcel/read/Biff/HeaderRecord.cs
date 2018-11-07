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
using common;
using NExcel;
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	/// <summary> A workbook page header record</summary>
	public class HeaderRecord:RecordData
	{
		/// <summary> Gets the header string
		/// 
		/// </summary>
		/// <returns> the header string
		/// </returns>
		virtual internal string Header
		{
			get
			{
				return header;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The footer</summary>
		private string header;
		
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
		internal HeaderRecord(Record t, WorkbookSettings ws):base(t)
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
				header = StringHelper.getUnicodeString(data, chars, 3);
			}
			else
			{
				header = StringHelper.getString(data, chars, 3, ws);
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
		internal HeaderRecord(Record t, WorkbookSettings ws, Biff7 dummy):base(t)
		{
			sbyte[] data = getRecord().Data;
			
			if (data.Length == 0)
			{
				return ;
			}
			
			int chars = data[0];
			header = StringHelper.getString(data, chars, 1, ws);
		}
		static HeaderRecord()
		{
			logger = Logger.getLogger(typeof(HeaderRecord));
			biff7 = new Biff7();
		}
	}
}
