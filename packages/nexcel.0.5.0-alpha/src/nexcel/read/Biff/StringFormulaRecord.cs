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
	
	/// <summary> A string formula's last calculated value</summary>
	class StringFormulaRecord:CellValue, LabelCell, FormulaData, StringFormulaCell
	{
		/// <summary> Interface method which returns the value
		/// 
		/// </summary>
		/// <returns> the last calculated value of the formula
		/// </returns>
		virtual public string Contents
		{
			get
			{
				return _Value;
			}
			
		}
		/// <summary> Interface method which returns the value
		/// 
		/// </summary>
		/// <returns> the last calculated value of the formula
		/// </returns>
		virtual public string StringValue
		{
			get
			{
				return _Value;
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
				return CellType.STRING_FORMULA;
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
		/// <summary> The last calculated value of the formula</summary>
		private string _Value;
		
		/// <summary> A handle to the class needed to access external sheets</summary>
		private ExternalSheet externalSheet;
		
		/// <summary> A handle to the name table</summary>
		private WorkbookMethods nameTable;
		
		/// <summary> The formula as an excel string</summary>
		private string formulaString;
		
		/// <summary> The raw data</summary>
		private sbyte[] data;
		
		/// <summary> Constructs this object from the raw data.  We need to use the excelFile
		/// to retrieve the String record which follows this formula record
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="excelFile">the excel file
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="es">the external sheet records
		/// </param>
		/// <param name="nt">the workbook
		/// </param>
		/// <param name="si">the sheet impl
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public StringFormulaRecord(Record t, File excelFile, FormattingRecords fr, ExternalSheet es, WorkbookMethods nt, SheetImpl si, WorkbookSettings ws):base(t, fr, si)
		{
			
			externalSheet = es;
			nameTable = nt;
			
			data = getRecord().Data;
			
			int pos = excelFile.Pos;
			
			// Look for the string record in one of the records after the
			// formula.  Put a cap on it to prevent looping
			
			Record nextRecord = excelFile.next();
			int count = 0;
			while (nextRecord.Type != NExcel.Biff.Type.STRING && count < 4)
			{
				nextRecord = excelFile.next();
				count++;
			}
			Assert.verify(count < 4, " @ " + pos);
			readString(nextRecord.Data, ws);
		}
		
		/// <summary> Reads in the string
		/// 
		/// </summary>
		/// <param name="d">the data
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		private void  readString(sbyte[] d, WorkbookSettings ws)
		{
			int pos = 0;
			int chars = IntegerHelper.getInt(d[0], d[1]);
			pos += 2;
			int optionFlags = d[pos];
			pos++;
			
			if ((optionFlags & 0xf) != optionFlags)
			{
				// Uh oh - looks like a plain old string, not unicode
				// Recalculate all the positions
				pos = 0;
				chars = IntegerHelper.getInt(d[0], (sbyte) 0);
				optionFlags = d[1];
				pos = 2;
			}
			
			// See if it is an extended string
			bool extendedString = ((optionFlags & 0x04) != 0);
			
			// See if string contains formatting information
			bool richString = ((optionFlags & 0x08) != 0);
			
			if (richString)
			{
				pos += 2;
			}
			
			if (extendedString)
			{
				pos += 4;
			}
			
			// See if string is ASCII (compressed) or unicode
			bool asciiEncoding = ((optionFlags & 0x01) == 0);
			
			//    byte[] bytes = null;
			
			if (asciiEncoding)
			{
				_Value = StringHelper.getString(d, chars, pos, ws);
			}
			else
			{
				_Value = StringHelper.getUnicodeString(d, chars, pos);
			}
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
