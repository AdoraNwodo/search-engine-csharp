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
using NExcel.Biff;
using NExcel.Biff.Formula;
namespace NExcel.Read.Biff
{
	
	/// <summary> A base class for shared formula records</summary>
	public abstract class BaseSharedFormulaRecord:CellValue, FormulaData
	{
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
					FormulaParser fp = new FormulaParser(tokens, this, externalSheet, nameTable, Sheet.Workbook.Settings);
					fp.parse();
					formulaString = fp.Formula;
				}
				
				return formulaString;
			}
			
		}
		/// <summary> Access for the external sheet
		/// 
		/// </summary>
		/// <returns> the external sheet
		/// </returns>
		virtual protected internal ExternalSheet ExternalSheet
		{
			get
			{
				return externalSheet;
			}
			
		}
		/// <summary> Access for the name table
		/// 
		/// </summary>
		/// <returns> the name table
		/// </returns>
		virtual protected internal WorkbookMethods NameTable
		{
			get
			{
				return nameTable;
			}
			
		}

		virtual public sbyte[] getFormulaData()
		{
			return null;
		}
			
		
		/// <summary> Accessor for the position of the next record
		/// 
		/// </summary>
		/// <returns> the position of the next record
		/// </returns>
		virtual internal int FilePos
		{
			get
			{
				return filePos;
			}
			
		}
		/// <summary> The formula as an excel string</summary>
		private string formulaString;
		
		/// <summary> The position of the next record in the file.  Used when looking for
		/// for subsequent records eg. a string value
		/// </summary>
		private int filePos;
		
		/// <summary> A handle to the formatting records</summary>
		new private FormattingRecords formattingRecords;
		
		/// <summary> The array of parsed tokens</summary>
		private sbyte[] tokens;
		
		/// <summary> The external sheet</summary>
		private ExternalSheet externalSheet;
		
		/// <summary> The name table</summary>
		private WorkbookMethods nameTable;
		
		/// <summary> Constructs this number
		/// 
		/// </summary>
		/// <param name="t">the record
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="es">the external sheet
		/// </param>
		/// <param name="nt">the name table
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		/// <param name="pos">the position of the next record in the file
		/// </param>
		public BaseSharedFormulaRecord(Record t, FormattingRecords fr, ExternalSheet es, WorkbookMethods nt, SheetImpl si, int pos):base(t, fr, si)
		{
			externalSheet = es;
			nameTable = nt;
			filePos = pos;
		}
		
		/// <summary> Called by the shared formula record to set the tokens for
		/// this formula
		/// 
		/// </summary>
		/// <param name="t">the tokens
		/// </param>
		internal virtual void  setTokens(sbyte[] t)
		{
			tokens = t;
		}
		
		/// <summary> Accessor for the tokens which make up this formula
		/// 
		/// </summary>
		/// <returns> the tokens
		/// </returns>
		protected internal sbyte[] getTokens()
		{
			return tokens;
		}
		
		/// <summary> In case the shared formula is not added for any reason, we need
		/// to expose the raw record data , in order to try again
		/// 
		/// </summary>
		/// <returns> the record data from the base class
		/// </returns>
		public override Record getRecord()
		{
			return base.getRecord();
		}
	}
}
