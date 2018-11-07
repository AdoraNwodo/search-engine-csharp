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
	public class SharedDateFormulaRecord:BaseSharedFormulaRecord, DateCell, FormulaData
	{
		/// <summary> Accessor for the contents as a string
		/// 
		/// </summary>
		/// <returns> the value as a string
		/// </returns>
		virtual public string Contents
		{
			get
			{
				return dateRecord.Contents;
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
				return CellType.DATE_FORMULA;
			}
			
		}
		/// <summary> Gets the date
		/// 
		/// </summary>
		/// <returns> the date
		/// </returns>
		virtual public DateTime DateValue
		{
			get
			{
				return dateRecord.DateValue;
			}
			
		}
		/// <summary> Indicates whether the date value contained in this cell refers to a date,
		/// or merely a time
		/// 
		/// </summary>
		/// <returns> TRUE if the value refers to a time
		/// </returns>
		virtual public bool Time
		{
			get
			{
				return dateRecord.Time;
			}
			
		}
		/// <summary> Gets the DateTimeFormatInfo used to format the cell.  This will normally be
		/// the format specified in the excel spreadsheet, but in the event of any
		/// difficulty parsing this, it will revert to the default date/time format.
		/// 
		/// </summary>
		/// <returns> the DateFormat object used to format the date in the original
		/// excel cell
		/// </returns>
		virtual public DateTimeFormatInfo DateFormat
		{
			get
			{
				return dateRecord.DateFormat;
			}
			
		}
		/// <summary> Re-use the date record to handle all the formatting information and
		/// date calculations
		/// </summary>
		private DateRecord dateRecord;
		
		/// <summary> The double value</summary>
		private double Value;
		
		/// <summary> Constructs this number formula
		/// 
		/// </summary>
		/// <param name="nfr">the number formula records
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="nf">flag indicating whether this uses the 1904 date system
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		/// <param name="pos">the position
		/// </param>
		public SharedDateFormulaRecord(SharedNumberFormulaRecord nfr, FormattingRecords fr, bool nf, SheetImpl si, int pos):base(nfr.getRecord(), fr, nfr.ExternalSheet, nfr.NameTable, si, pos)
		{
			dateRecord = new DateRecord(nfr, nfr.XFIndex, fr, nf, si);
			Value = nfr.DoubleValue;
		}
		
		/// <summary> Accessor for the value
		/// 
		/// </summary>
		/// <returns> the value
		/// </returns>
		public virtual double getValue()
		{
			return Value;
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
			DoubleHelper.getIEEEBytes(Value, data, 6);
			
			// Now copy in the parsed tokens
			Array.Copy(rpnTokens, 0, data, 22, rpnTokens.Length);
			IntegerHelper.getTwoBytes(rpnTokens.Length, data, 20);
			
			// Lop off the standard information
			sbyte[] d = new sbyte[data.Length - 6];
			Array.Copy(data, 6, d, 0, data.Length - 6);
			
			return d;
		}
	}
}
