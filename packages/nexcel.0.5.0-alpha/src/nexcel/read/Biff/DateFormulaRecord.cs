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
using NExcelUtils;
using NExcel;
using NExcel.Biff;
using NExcel.Biff.Formula;
namespace NExcel.Read.Biff
{
	
	/// <summary> A date formula's last calculated value</summary>
	class DateFormulaRecord:DateRecord, DateCell, FormulaData, DateFormulaCell
	{
		/// <summary> Returns the cell type
		/// 
		/// </summary>
		/// <returns> The cell type
		/// </returns>
		override public CellType Type
		{
			get
			{
				return CellType.DATE_FORMULA;
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
				// Note that the standard information was lopped off by the NumberFormula
				// record when creating this formula
				if ((System.Object) formulaString == null)
				{
					sbyte[] tokens = new sbyte[data.Length - 16];
					Array.Copy(data, 16, tokens, 0, tokens.Length);
					FormulaParser fp = new FormulaParser(tokens, this, externalSheet, nameTable, Sheet.Workbook.Settings);
					fp.parse();
					formulaString = fp.Formula;
				}
				
				return formulaString;
			}
			
		}
		/// <summary> Dummy implementation in order to adhere to the NumberCell interface
		/// 
		/// </summary>
		/// <returns> NULL
		/// </returns>
		virtual public NumberFormatInfo NumberFormat
		{
			get
			{
				return null;
			}
			
		}
		/// <summary> The last calculated value of the formula</summary>
		private double Value;
		
		/// <summary> The formula as an excel string</summary>
		private string formulaString;
		
		/// <summary> A handle to the class needed to access external sheets</summary>
		private ExternalSheet externalSheet;
		
		/// <summary> A handle to the name table</summary>
		private WorkbookMethods nameTable;
		
		/// <summary> The raw data</summary>
		private sbyte[] data;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the basic number formula record
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="es">the external sheet
		/// </param>
		/// <param name="nt">the name table
		/// </param>
		/// <param name="nf">flag indicating whether the 1904 date system is in use
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		public DateFormulaRecord(NumberFormulaRecord t, FormattingRecords fr, ExternalSheet es, WorkbookMethods nt, bool nf, SheetImpl si):base(t, t.XFIndex, fr, nf, si)
		{
			
			externalSheet = es;
			nameTable = nt;
			data = t.getFormulaData();
		}
		
		/// <summary> Gets the raw bytes for the formula.  This will include the
		/// parsed tokens array.  Used when copying spreadsheets
		/// 
		/// </summary>
		/// <returns> the raw record data
		/// </returns>
		public virtual sbyte[] getFormulaData()
		{
			// Data is already the formula data, so don't do any more manipulation
			return data;
		}
		
		/// <summary> Interface method which returns the value
		/// 
		/// </summary>
		/// <returns> the last calculated value of the formula
		/// </returns>
		public virtual double getValue()
		{
			return Value;
		}
	}
}
