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
using NExcel.Read.Biff;
namespace NExcel.Biff
{
	
	/// <summary> The record data within a record</summary>
	public abstract class RecordData 
	{
		/// <summary> Accessor for the code
		/// 
		/// </summary>
		/// <returns> the code
		/// </returns>
		virtual protected internal int Code
		{
			get
			{
				return code;
			}
			
		}
		/// <summary> The raw data</summary>
		private Record record;
		
		/// <summary> The Biff code for this record.  This is set up when the record is
		/// used for writing
		/// </summary>
		private int code;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="r">the raw data
		/// </param>
		protected internal RecordData(Record r)
		{
			record = r;
			code = r.Code;
		}
		
		/// <summary> Constructor used by the writable records
		/// 
		/// </summary>
		/// <param name="t">the type
		/// </param>
		protected internal RecordData(NExcel.Biff.Type t)
		{
			code = t.Value;
		}
		
		/// <summary> Returns the raw data to its subclasses
		/// 
		/// </summary>
		/// <returns> the raw data
		/// </returns>
		public virtual Record getRecord()
		{
			return record;
		}
	}
}
