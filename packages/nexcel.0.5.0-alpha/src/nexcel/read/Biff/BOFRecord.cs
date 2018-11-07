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
	
	/// <summary> A Beginning Of File record, found at the commencement of all substreams
	/// within a biff8 file
	/// </summary>
	public class BOFRecord:RecordData
	{
		/// <summary> Gets the .Length of the data portion of this record
		/// Used to adjust when reading sheets which contain just a chart
		/// </summary>
		/// <returns> the .Length of the data portion of this record
		/// </returns>
		virtual internal int Length
		{
			get
			{
				return getRecord().Length;
			}
			
		}
		/// <summary> The code used for biff8 files</summary>
		private const int Biff8 = 0x600;
		/// <summary> The code used for biff8 files</summary>
		private const int Biff7 = 0x500;
		/// <summary> The code used for workbook globals</summary>
		private const int WorkbookGlobals = 0x5;
		/// <summary> The code used for worksheets</summary>
		private const int Worksheet = 0x10;
		/// <summary> The code used for charts</summary>
		private const int Chart = 0x20;
		/// <summary> The code used for macro sheets</summary>
		private const int MacroSheet = 0x40;
		
		/// <summary> The biff version of this substream</summary>
		private int version;
		/// <summary> The type of this substream</summary>
		private int substreamType;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		internal BOFRecord(Record t):base(t)
		{
			sbyte[] data = getRecord().Data;
			version = IntegerHelper.getInt(data[0], data[1]);
			substreamType = IntegerHelper.getInt(data[2], data[3]);
		}
		
		/// <summary> Interrogates this object to see if it is a biff8 substream
		/// 
		/// </summary>
		/// <returns> TRUE if this substream is biff8, false otherwise
		/// </returns>
		public virtual bool isBiff8()
		{
			return version == Biff8;
		}
		
		/// <summary> Interrogates this object to see if it is a biff7 substream
		/// 
		/// </summary>
		/// <returns> TRUE if this substream is biff7, false otherwise
		/// </returns>
		public virtual bool isBiff7()
		{
			return version == Biff7;
		}
		
		
		/// <summary> Interrogates this substream to see if it represents the commencement of
		/// the workbook globals substream
		/// 
		/// </summary>
		/// <returns> TRUE if this is the commencement of a workbook globals substream,
		/// FALSE otherwise
		/// </returns>
		internal virtual bool isWorkbookGlobals()
		{
			return substreamType == WorkbookGlobals;
		}
		
		/// <summary> Interrogates the substream to see if it is the commencement of a worksheet
		/// 
		/// </summary>
		/// <returns> TRUE if this substream is the beginning of a worksheet, FALSE
		/// otherwise
		/// </returns>
		public virtual bool isWorksheet()
		{
			return substreamType == Worksheet;
		}
		
		/// <summary> Interrogates the substream to see if it is the commencement of a worksheet
		/// 
		/// </summary>
		/// <returns> TRUE if this substream is the beginning of a worksheet, FALSE
		/// otherwise
		/// </returns>
		public virtual bool isMacroSheet()
		{
			return substreamType == MacroSheet;
		}
		
		/// <summary> Interrogates the substream to see if it is a chart
		/// 
		/// </summary>
		/// <returns> TRUE if this substream is the beginning of a worksheet, FALSE
		/// otherwise
		/// </returns>
		public virtual bool isChart()
		{
			return substreamType == Chart;
		}
	}
}
