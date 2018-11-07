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
	
	/// <summary> A record which merely holds the MSODRAWING data.  Used when copying files
	/// which contain images
	/// </summary>
	public class MsoDrawingRecord:WritableRecordData
	{
		/// <summary> The raw drawing data which was read in</summary>
		private sbyte[] data;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public MsoDrawingRecord(Record t):base(t)
		{
			data = getRecord().Data;
		}
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="d">the drawing data
		/// </param>
		public MsoDrawingRecord(sbyte[] d):base(NExcel.Biff.Type.MSODRAWING)
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
		
		/// <summary> Expose the protected function to the SheetImpl in this package
		/// 
		/// </summary>
		/// <returns> the raw record data
		/// </returns>
		public override Record getRecord()
		{
			return base.getRecord();
		}
	}
}
