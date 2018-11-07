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
	
	/// <summary> A boolean cell last calculated value</summary>
	class BooleanRecord:CellValue, BooleanCell
	{
		/// <summary> Interface method which queries whether this cell contains an error.
		/// Returns TRUE if it does, otherwise returns FALSE.
		/// 
		/// </summary>
		/// <returns> TRUE if this cell is an error, FALSE otherwise
		/// </returns>
		virtual public bool Error
		{
			get
			{
				return error;
			}
			
		}
		/// <summary> Interface method which Gets the boolean value stored in this cell.  If
		/// this cell contains an error, then returns FALSE.  Always query this cell
		/// type using the accessor method isError() prior to calling this method
		/// 
		/// </summary>
		/// <returns> TRUE if this cell contains TRUE, FALSE if it contains FALSE or
		/// an error code
		/// </returns>
		virtual public bool BooleanValue
		{
			get
			{
				return _Value;
			}
		}

		/// <summary>
		/// Returns the value.
		/// </summary>

		public virtual object Value
		{
			get
			{
				return this._Value;
			}
		}

		/// <summary> Returns the numerical value as a string
		/// 
		/// </summary>
		/// <returns> The numerical value of the formula as a string
		/// </returns>
		virtual public string Contents
		{
			get
			{
				Assert.verify(!Error);
				
				// [TODO-NExcel_Next] - check if it is right in different languages
				//return Boolean.toString(_Value);
				return _Value.ToString().ToUpper();
			}
			
		}
		/// <summary> Returns the cell type
		/// 
		/// </summary>
		/// <returns> The cell type
		/// </returns>
		virtual public CellType Type
		{
			get
			{
				return CellType.BOOLEAN;
			}
			
		}
		/// <summary> Indicates whether this cell contains an error or a boolean</summary>
		private bool error;
		
		/// <summary> The boolean value of this cell.  If this cell represents an error,
		/// this will be false
		/// </summary>
		private bool _Value;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="fr"> the formatting records
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		public BooleanRecord(Record t, FormattingRecords fr, SheetImpl si):base(t, fr, si)
		{
			error = false;
			_Value = false;
			
			sbyte[] data = getRecord().Data;
			
			error = (data[7] == 1);
			
			if (!error)
			{
				_Value = data[6] == 1?true:false;
			}
		}
		
		/// <summary> A special case which overrides the method in the subclass to get
		/// hold of the raw data
		/// 
		/// </summary>
		/// <returns> the record
		/// </returns>
		public override Record getRecord()
		{
			return base.getRecord();
		}
	}
}
