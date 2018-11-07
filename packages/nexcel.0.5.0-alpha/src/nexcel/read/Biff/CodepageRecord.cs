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
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	class CodepageRecord:RecordData
	{
		/// <summary> Accessor for the encoding
		/// 
		/// </summary>
		/// <returns> the character encoding
		/// </returns>
		virtual public int CharacterSet
		{
			get
			{
				return characterSet;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The character encoding</summary>
		private int characterSet;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="t">the record
		/// </param>
		public CodepageRecord(Record t):base(t)
		{
			sbyte[] data = t.Data;
			characterSet = IntegerHelper.getInt(data[0], data[1]);
		}
		static CodepageRecord()
		{
			logger = Logger.getLogger(typeof(CodepageRecord));
		}
	}
}
