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
	
	/// <summary> A cell containing an error code.  This will usually be the result
	/// of some error during the calculation of a formula
	/// </summary>
	class ErrorRecord:CellValue, ErrorCell
	{
		/// <summary> Interface method which gets the error code for this cell.  If this cell
		/// does not represent an error, then it returns 0.  Always use the
		/// method isError() to  determine this prior to calling this method
		/// 
		/// </summary>
		/// <returns> the error code if this cell contains an error, 0 otherwise
		/// </returns>
		virtual public int ErrorCode
		{
			get
			{
				return errorCode;
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
				return "ERROR " + errorCode;
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
				return CellType.ERROR;
			}
			
		}
		/// <summary> The error code if this cell evaluates to an error, otherwise zer0</summary>
		private int errorCode;
		
		/// <summary> Constructs this object
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		public ErrorRecord(Record t, FormattingRecords fr, SheetImpl si):base(t, fr, si)
		{
			
			sbyte[] data = getRecord().Data;
			
			errorCode = data[6];
		}
	}
}
