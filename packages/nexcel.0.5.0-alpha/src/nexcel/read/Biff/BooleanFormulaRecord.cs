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
using NExcel.Biff.Formula;
namespace NExcel.Read.Biff
{
	
	/// <summary> A boolean formula's last calculated value</summary>
	class BooleanFormulaRecord:CellValue, BooleanCell, FormulaData, BooleanFormulaCell
	{
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
				return CellType.BOOLEAN_FORMULA;
			}
			
		}
		/// <summary> Gets the formula as an excel string
		/// 
		/// </summary>
		/// <returns> the formula as an excel string
		/// </returns>
		/// <exception cref=""> FormulaException
		/// </exception>
		virtual public string Formula
		{
			get
			{
				if ((System.Object) formulaString == null)
				{
					sbyte[] tokens = new sbyte[data.Length - 22];
					Array.Copy(data, 22, tokens, 0, tokens.Length);
					FormulaParser fp = new FormulaParser(tokens, this, externalSheet, nameTable, Sheet.Workbook.Settings);
					fp.parse();
					formulaString = fp.Formula;
				}
				
				return formulaString;
			}
			
		}
		/// <summary> The boolean value of this cell.  If this cell represents an error,
		/// this will be false
		/// </summary>
		private bool _Value;
		
		/// <summary> A handle to the class needed to access external sheets</summary>
		private ExternalSheet externalSheet;
		
		/// <summary> A handle to the name table</summary>
		private WorkbookMethods nameTable;
		
		/// <summary> The formula as an excel string</summary>
		private string formulaString;
		
		/// <summary> The raw data</summary>
		private sbyte[] data;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		/// <param name="es">the sheet
		/// </param>
		/// <param name="nt">the name table
		/// </param>
		public BooleanFormulaRecord(Record t, FormattingRecords fr, ExternalSheet es, WorkbookMethods nt, SheetImpl si):base(t, fr, si)
		{
			externalSheet = es;
			nameTable = nt;
			_Value = false;
			
			data = getRecord().Data;
			
			Assert.verify(data[6] != 2);
			
			_Value = data[8] == 1?true:false;
		}
		
		/// <summary> Gets the raw bytes for the formula.  This will include the
		/// parsed tokens array
		/// 
		/// </summary>
		/// <returns> the raw record data
		/// </returns>
		public virtual sbyte[] getFormulaData()
		{
			// Lop off the standard information
			sbyte[] d = new sbyte[data.Length - 6];
			Array.Copy(data, 6, d, 0, data.Length - 6);
			
			return d;
		}
	}
}
