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
	
	/// <summary> A string formula record, manufactured out of the Shared Formula
	/// "optimization"
	/// </summary>
	public class SharedStringFormulaRecord:BaseSharedFormulaRecord, LabelCell, FormulaData, StringFormulaCell
	{
		/// <summary> Accessor for the value
		/// 
		/// </summary>
		/// <returns> the value
		/// </returns>
		virtual public string StringValue
		{
			get
			{
				return _Value;
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
				return CellType.STRING_FORMULA;
			}
			
		}
		/// <summary> The value of this string formula</summary>
		private string _Value;
		/// <summary> A handle to the formatting records</summary>
		new private FormattingRecords formattingRecords;
		
		/// <summary> Constructs this string formula
		/// 
		/// </summary>
		/// <param name="t">the record
		/// </param>
		/// <param name="excelFile">the excel file
		/// </param>
		/// <param name="fr">the formatting record
		/// </param>
		/// <param name="es">the external sheet
		/// </param>
		/// <param name="nt">the workbook
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public SharedStringFormulaRecord(Record t, File excelFile, FormattingRecords fr, ExternalSheet es, WorkbookMethods nt, SheetImpl si, WorkbookSettings ws):base(t, fr, es, nt, si, excelFile.Pos)
		{
			int pos = excelFile.Pos;
			
			// Save the position in the excel file
			int filepos = excelFile.Pos;
			
			// Look for the string record in one of the records after the
			// formula.  Put a cap on it to prevent ednas
			Record nextRecord = excelFile.next();
			int count = 0;
			while (nextRecord.Type != NExcel.Biff.Type.STRING && count < 4)
			{
				nextRecord = excelFile.next();
				count++;
			}
			Assert.verify(count < 4, " @ " + pos);
			
			sbyte[] stringData = nextRecord.Data;
			int chars = IntegerHelper.getInt(stringData[0], stringData[1]);
			
			bool unicode = false;
			int startpos = 3;
			if (stringData.Length == chars + 2)
			{
				// String might only consist of a one byte .Length indicator, instead
				// of the more normal 2
				startpos = 2;
				unicode = false;
			}
			else if (stringData[2] == 0x1)
			{
				// unicode string, two byte .Length indicator
				startpos = 3;
				unicode = true;
			}
			else
			{
				// ascii string, two byte .Length indicator
				startpos = 3;
				unicode = false;
			}
			
			if (!unicode)
			{
				_Value = StringHelper.getString(stringData, chars, startpos, ws);
			}
			else
			{
				_Value = StringHelper.getUnicodeString(stringData, chars, startpos);
			}
			
			// Restore the position in the excel file, to enable the SHRFMLA
			// record to be picked up
			excelFile.Pos = filepos;
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
			
			// Set the two most significant bytes of the value to be 0xff in
			// order to identify this as a string
			data[6] = 0;
			data[12] = - 1;
			data[13] = - 1;
			
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
