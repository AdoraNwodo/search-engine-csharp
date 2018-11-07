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
using NExcel.Read.Biff;
namespace NExcel.Biff.Drawing
{
	
	/// <summary> A record which merely holds the MSODRAWINGGROUP data.  Used when copying
	/// files  which contain images
	/// </summary>
	public class MsoDrawingGroupRecord:WritableRecordData
	{
		private sbyte[] data;
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public MsoDrawingGroupRecord(Record t):base(t)
		{
			data = t.Data;
		}
		
		internal MsoDrawingGroupRecord(sbyte[] d):base(NExcel.Biff.Type.MSODRAWINGGROUP)
		{
			data = d;
		}
		
		/// <summary> Expose the protected function to the SheetImpl in this package
		/// 
		/// </summary>
		/// <returns> the raw record data
		/// </returns>
		public override sbyte[] getData()
		{
			return data;
		}
	}
}
