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
	
	/// <summary> A number formula record, manufactured out of the Shared Formula
	/// "optimization"
	/// </summary>
	public class SharedNumberFormulaRecord:BaseSharedFormulaRecord, NumberCell, FormulaData, NumberFormulaCell
	{
		/// <summary> Accessor for the value
		/// 
		/// </summary>
		/// <returns> the value
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
		public override object Value
		{
			get
			{
				return this._Value;
			}
		}

		/// <summary> Accessor for the contents as a string
		/// 
		/// </summary>
		/// <returns> the value as a string
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
		/// <summary> Accessor for the cell type
		/// 
		/// </summary>
		/// <returns> the cell type
		/// </returns>
		virtual public CellType Type
		{
			get
			{
				return CellType.NUMBER_FORMULA;
			}
			
		}
		/// <summary> The value of this number</summary>
		private double _Value;
		/// <summary> The cell format</summary>
		new private NumberFormatInfo format;
		/// <summary> A handle to the formatting records</summary>
		new private FormattingRecords formattingRecords;
		
		/// <summary> The string format for the double value</summary>
		private static NumberFormatInfo defaultFormat;
		
		/// <summary> Constructs this number
		/// 
		/// </summary>
		/// <param name="t">the data
		/// </param>
		/// <param name="excelFile">the excel biff data
		/// </param>
		/// <param name="v">the value
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="es">the external sheet
		/// </param>
		/// <param name="nt">the name table
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		public SharedNumberFormulaRecord(Record t, File excelFile, double v, FormattingRecords fr, ExternalSheet es, WorkbookMethods nt, SheetImpl si):base(t, fr, es, nt, si, excelFile.Pos)
		{
			_Value = v;
			format = defaultFormat;
		}
		
		/// <summary> Sets the format for the number based on the Excel spreadsheets' format.
		/// This is called from SheetImpl when it has been definitely established
		/// that this cell is a number and not a date
		/// 
		/// </summary>
		/// <param name="f">the format
		/// </param>
		internal void  setNumberFormat(NumberFormatInfo f)
		{
			if (f != null)
			{
				format = f;
			}
		}
		
		/// <summary> Gets the raw bytes for the formula.  This will include the
		/// parsed tokens array.  Used when copying spreadsheets
		/// 
		/// </summary>
		/// <returns> the raw record data
		/// </returns>
		/// <exception cref=""> FormulaException
		/// </exception>
		public override sbyte[] getFormulaData()
		{
			// Get the tokens, taking into account the mapping from shared
			// formula specific values into normal values
			FormulaParser fp = new FormulaParser(getTokens(), this, ExternalSheet, NameTable, Sheet.Workbook.Settings);
			fp.parse();
			sbyte[] rpnTokens = fp.Bytes;
			
			sbyte[] data = new sbyte[rpnTokens.Length + 22];
			
			// Set the standard info for this cell
			IntegerHelper.getTwoBytes(Row, data, 0);
			IntegerHelper.getTwoBytes(Column, data, 2);
			IntegerHelper.getTwoBytes(XFIndex, data, 4);
			DoubleHelper.getIEEEBytes(_Value, data, 6);
			
			// Now copy in the parsed tokens
			Array.Copy(rpnTokens, 0, data, 22, rpnTokens.Length);
			IntegerHelper.getTwoBytes(rpnTokens.Length, data, 20);
			
			// Lop off the standard information
			sbyte[] d = new sbyte[data.Length - 6];
			Array.Copy(data, 6, d, 0, data.Length - 6);
			
			return d;
		}
		
		/// <summary> Gets the NumberFormatInfo used to format this cell.  This is the java
		/// equivalent of the Excel format
		/// 
		/// </summary>
		/// <returns> the NumberFormatInfo used to format the cell
		/// </returns>
		public virtual NumberFormatInfo NumberFormat
		{
		get
		{
		return format;
		}
		}
		static SharedNumberFormulaRecord()
		{
			defaultFormat = new NumberFormatInfo("#.###");
		}
	}
}
