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
	
	/// <summary> An Externsheet record, containing the details of externally references
	/// workbooks
	/// </summary>
	public class ExternalSheetRecord:RecordData
	{
		/// <summary> Accessor for  the number of external sheet records</summary>
		/// <returns> the number of XTI records
		/// </returns>
		virtual public int NumRecords
		{
			get
			{
				return xtiArray.Length;
			}
			
		}
		/// <summary> Used when copying a workbook to access the raw external sheet data
		/// 
		/// </summary>
		/// <returns> the raw external sheet data
		/// </returns>
		virtual public sbyte[] Data
		{
			get
			{
				return getRecord().Data;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> Dummy indicators for overloading the constructor</summary>
		public class Biff7
		{
		}
		
		public static Biff7 biff7;
		
		/// <summary> An XTI structure</summary>
		private class XTI
		{
			/// <summary> the supbook index</summary>
			internal int supbookIndex;
			/// <summary> the first tab</summary>
			internal int firstTab;
			/// <summary> the last tab</summary>
			internal int lastTab;
			
			/// <summary> Constructor
			/// 
			/// </summary>
			/// <param name="s">the supbook index
			/// </param>
			/// <param name="f">the first tab
			/// </param>
			/// <param name="l">the last tab
			/// </param>
			internal XTI(int s, int f, int l)
			{
				supbookIndex = s;
				firstTab = f;
				lastTab = l;
			}
		}
		
		/// <summary> The array of XTI structures</summary>
		private XTI[] xtiArray;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		internal ExternalSheetRecord(Record t, WorkbookSettings ws):base(t)
		{
			sbyte[] data = getRecord().Data;
			
			int numxtis = IntegerHelper.getInt(data[0], data[1]);
			
			if (data.Length < numxtis * 6 + 2)
			{
				xtiArray = new XTI[0];
				logger.warn("Could not process external sheets.  Formulas may be compromised.");
				return ;
			}
			
			xtiArray = new XTI[numxtis];
			
			int pos = 2;
			for (int i = 0; i < numxtis; i++)
			{
				int s = IntegerHelper.getInt(data[pos], data[pos + 1]);
				int f = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
				int l = IntegerHelper.getInt(data[pos + 4], data[pos + 5]);
				xtiArray[i] = new XTI(s, f, l);
				pos += 6;
			}
		}
		
		/// <summary> Constructs this object from the raw data in biff 7 format.
		/// Does nothing here
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="settings">the workbook settings
		/// </param>
		/// <param name="dummy">dummy override to identify biff7 funcionality
		/// </param>
		internal ExternalSheetRecord(Record t, WorkbookSettings settings, Biff7 dummy):base(t)
		{
		}
		/// <summary> Gets the supbook index for the specified external sheet
		/// 
		/// </summary>
		/// <param name="index">the index of the supbook record
		/// </param>
		/// <returns> the supbook index
		/// </returns>
		public virtual int getSupbookIndex(int index)
		{
			return xtiArray[index].supbookIndex;
		}
		
		/// <summary> Gets the first tab index for the specified external sheet
		/// 
		/// </summary>
		/// <param name="the">index of the supbook record
		/// </param>
		/// <returns> the first tab index
		/// </returns>
		public virtual int getFirstTabIndex(int index)
		{
			return xtiArray[index].firstTab;
		}
		
		/// <summary> Gets the last tab index for the specified external sheet
		/// 
		/// </summary>
		/// <param name="index">the index of the supbook record
		/// </param>
		/// <returns> the last tab index
		/// </returns>
		public virtual int getLastTabIndex(int index)
		{
			return xtiArray[index].lastTab;
		}
		static ExternalSheetRecord()
		{
			logger = Logger.getLogger(typeof(ExternalSheetRecord));
			biff7 = new Biff7();
		}
	}
}
