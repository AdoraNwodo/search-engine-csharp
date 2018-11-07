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
	
	/// <summary> A label which is stored in the shared string table</summary>
	class LabelSSTRecord:CellValue, LabelCell
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
		override public object Value
		{
			get
			{
				return this._Value;
			}
		}

		/// <summary> Gets this cell's contents as a string
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
		/// <summary> Returns the cell type
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
		/// <summary> The index into the shared string table</summary>
		private int index;
		/// <summary> The label</summary>
		private string _Value;
		
		/// <summary> Constructor.  Retrieves the index from the raw data and looks it up
		/// in the shared string table
		/// 
		/// </summary>
		/// <param name="stringTable">the shared string table
		/// </param>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		public LabelSSTRecord(Record t, SSTRecord stringTable, FormattingRecords fr, SheetImpl si):base(t, fr, si)
		{
			sbyte[] data = getRecord().Data;
			index = IntegerHelper.getInt(data[6], data[7], data[8], data[9]);
			_Value = stringTable.getString(index);
		}
	}
}
