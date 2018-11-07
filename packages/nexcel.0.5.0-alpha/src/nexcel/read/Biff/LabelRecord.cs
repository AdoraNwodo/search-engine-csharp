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
	
	/// <summary> A label which is stored in the cell</summary>
	class LabelRecord:CellValue, LabelCell
	{
		/// <summary> Gets the label
		/// 
		/// </summary>
		/// <returns> the label
		/// </returns>
		virtual public string StringValue
		{
			get
			{
				return _Value;
			}
			
		}

		/// <summary>
		/// Returns the string value.
		/// </summary>
		virtual public object Value
		{
			get
			{
				return this._Value;
			}
		}


		/// <summary> Gets the cell contents as a string
		/// 
		/// </summary>
		/// <returns> the label
		/// </returns>
		virtual public string Contents
		{
			get
			{
				return _Value;
			}
			
		}
		/// <summary> Accessor for the cell type
		/// 
		/// </summary>
		/// <returns> the cell type
		/// </returns>
		virtual public CellType Type
		{
			get
			{
				return CellType.LABEL;
			}
			
		}
		/// <summary> The length of the label in characters</summary>
		private int length;
		/// <summary> The label</summary>
		private string _Value;
		
		/// <summary> Dummy indicators for overloading the constructor</summary>
		public class Biff7
		{
		}
		
		public static Biff7 biff7;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public LabelRecord(Record t, FormattingRecords fr, SheetImpl si, WorkbookSettings ws):base(t, fr, si)
		{
			sbyte[] data = getRecord().Data;
			length = IntegerHelper.getInt(data[6], data[7]);
			
			if (data[8] == 0x0)
			{
				_Value = StringHelper.getString(data, length, 9, ws);
			}
			else
			{
				_Value = StringHelper.getUnicodeString(data, length, 9);
			}
		}
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <param name="dummy">dummy overload to indicate a biff 7 workbook
		/// </param>
		public LabelRecord(Record t, FormattingRecords fr, SheetImpl si, WorkbookSettings ws, Biff7 dummy):base(t, fr, si)
		{
			sbyte[] data = getRecord().Data;
			length = IntegerHelper.getInt(data[6], data[7]);
			
			_Value = StringHelper.getString(data, length, 8, ws);
		}
		static LabelRecord()
		{
			biff7 = new Biff7();
		}
	}
}
