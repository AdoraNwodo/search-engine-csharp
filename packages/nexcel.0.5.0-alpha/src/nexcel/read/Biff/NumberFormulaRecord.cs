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
using NExcel.Biff.Formula;
using NExcelUtils;

namespace NExcel.Read.Biff
{
	
	/// <summary> A formula's last calculated value</summary>
	class NumberFormulaRecord:CellValue, NumberCell, FormulaData, NumberFormulaCell
	{
		/// <summary> Interface method which returns the value
		/// 
		/// </summary>
		/// <returns> the last calculated value of the formula
		/// </returns>
		virtual public double DoubleValue
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
				// [TODO-NExcel_Next] find a better way
//				return _Value.ToString(format);
				return string.Format(format, "{0}", _Value);
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
				return CellType.NUMBER_FORMULA;
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
		/// <summary> Gets the NumberFormatInfo used to format this cell.  This is the java
		/// equivalent of the Excel format
		/// 
		/// </summary>
		/// <returns> the NumberFormatInfo used to format the cell
		/// </returns>
		virtual public NumberFormatInfo NumberFormat
		{
			get
			{
				return format;
			}
			
		}
		/// <summary> The last calculated value of the formula</summary>
		private double _Value;
		
		/// <summary> The number format</summary>
		new private NumberFormatInfo format;
		
		/// <summary> The string format for the double value</summary>
		private static readonly NumberFormatInfo defaultFormat = new NumberFormatInfo("#.###");
		
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
		/// <param name="t">the raw data
		/// </param>
		/// <param name="fr">the formatting record
		/// </param>
		/// <param name="es">the external sheet
		/// </param>
		/// <param name="nt">the name table
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		public NumberFormulaRecord(Record t, FormattingRecords fr, ExternalSheet es, WorkbookMethods nt, SheetImpl si):base(t, fr, si)
		{
			
			externalSheet = es;
			nameTable = nt;
			data = getRecord().Data;
			
			format = fr.getNumberFormat(XFIndex);
			
			if (format == null)
			{
				format = defaultFormat;
			}
			
			int num1 = IntegerHelper.getInt(data[6], data[7], data[8], data[9]);
			int num2 = IntegerHelper.getInt(data[10], data[11], data[12], data[13]);
			
			// bitwise ors don't work with longs, so we have to simulate this
			// functionality the long way round by concatenating two binary
			// strings, and then parsing the binary string into a long.
			// This is very clunky and inefficient, and I hope to
			// find a better way
			string s1 = System.Convert.ToString(num1, 2);
			while (s1.Length < 32)
			{
				s1 = "0" + s1; // fill out with leading zeros as necessary
			}
			
			// Long.parseLong doesn't like the sign bit, so have to extract this
			// information and put it in at the end.  (thanks
			// to Ruben for pointing this out)
			bool negative = ((((long) num2) & 0x80000000) != 0);
			
			string s = System.Convert.ToString(num2 & 0x7fffffff, 2) + s1;
			long val = System.Convert.ToInt64(s, 2);

			_Value = BitConverter.Int64BitsToDouble(val);
			
			if (negative)
			{
				_Value = - _Value;
			}
		}
		
		/// <summary> Gets the raw bytes for the formula.  This will include the
		/// parsed tokens array.  Used when copying spreadsheets
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
