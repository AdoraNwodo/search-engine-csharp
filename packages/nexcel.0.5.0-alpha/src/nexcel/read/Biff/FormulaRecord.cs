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
	
	/// <summary> A formula's last calculated value</summary>
	class FormulaRecord:CellValue
	{
		/// <summary> Returns the numerical value as a string
		/// 
		/// </summary>
		/// <returns> The numerical value of the formula as a string
		/// </returns>
		virtual public string Contents
		{
			get
			{
				Assert.verify(false);
				return "";
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
				Assert.verify(false);
				return CellType.EMPTY;
			}
			
		}
		/// <summary> Gets the "real" formula
		/// 
		/// </summary>
		/// <returns>  the cell value
		/// </returns>
		virtual internal CellValue Formula
		{
			get
			{
				return formula;
			}
			
		}
		/// <summary> Interrogates this formula to determine if it forms part of a shared
		/// formula
		/// 
		/// </summary>
		/// <returns> TRUE if this is shared formula, FALSE otherwise
		/// </returns>
		virtual internal bool Shared
		{
			get
			{
				return shared;
			}
			
		}
		/// <summary> The logger</summary>
		new private static Logger logger;
		
		/// <summary> The "real" formula record - will be either a string a or a number</summary>
		private CellValue formula;
		
		/// <summary> Flag to indicate whether this is a shared formula</summary>
		private bool shared;
		
		/// <summary> Static class for a dummy override, indicating that the formula
		/// passed in is not a shared formula
		/// </summary>
		public class IgnoreSharedFormula
		{
		}
		
		public static readonly IgnoreSharedFormula ignoreSharedFormula = new IgnoreSharedFormula();
		
		/// <summary> Constructs this object from the raw data.  Creates either a
		/// NumberFormulaRecord or a StringFormulaRecord depending on whether
		/// this formula represents a numerical calculation or not
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="excelFile">the excel file
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="es">the workbook, which contains the external sheet references
		/// </param>
		/// <param name="nt">the name table
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public FormulaRecord(Record t, File excelFile, FormattingRecords fr, ExternalSheet es, WorkbookMethods nt, SheetImpl si, WorkbookSettings ws):base(t, fr, si)
		{
			
			sbyte[] data = getRecord().Data;
			
			shared = false;
			
			// Check to see if this forms part of a shared formula
			int grbit = IntegerHelper.getInt(data[14], data[15]);
			if ((grbit & 0x08) != 0)
			{
				shared = true;
				
				if (data[6] == 0 && data[12] == - 1 && data[13] == - 1)
				{
					// It is a shared string formula
					formula = new SharedStringFormulaRecord(t, excelFile, fr, es, nt, si, ws);
				}
				else
				{
					// It is a numerical formula
					double Value = DoubleHelper.getIEEEDouble(data, 6);
					formula = new SharedNumberFormulaRecord(t, excelFile, Value, fr, es, nt, si);
				}
				
				return ;
			}
			
			// microsoft and their goddam magic values determine whether this
			// is a string or a number value
			if (data[6] == 0 && data[12] == - 1 && data[13] == - 1)
			{
				// we have a string
				formula = new StringFormulaRecord(t, excelFile, fr, es, nt, si, ws);
			}
			else if (data[6] == 1 && data[12] == - 1 && data[13] == - 1)
			{
				// We have a boolean formula
				// multiple values.  Thanks to Frank for spotting this
				formula = new BooleanFormulaRecord(t, fr, es, nt, si);
			}
			else if (data[6] == 2 && data[12] == - 1 && data[13] == - 1)
			{
				// The cell is in error
				formula = new ErrorFormulaRecord(t, fr, es, nt, si);
			}
			else
			{
				// it is most assuredly a number
				formula = new NumberFormulaRecord(t, fr, es, nt, si);
			}
		}
		
		/// <summary> Constructs this object from the raw data.  Creates either a
		/// NumberFormulaRecord or a StringFormulaRecord depending on whether
		/// this formula represents a numerical calculation or not
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="excelFile">the excel file
		/// </param>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="es">the workbook, which contains the external sheet references
		/// </param>
		/// <param name="nt">the name table
		/// </param>
		/// <param name="i">a dummy override to indicate that we don't want to do
		/// any shared formula processing
		/// </param>
		/// <param name="si">the sheet impl
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public FormulaRecord(Record t, File excelFile, FormattingRecords fr, ExternalSheet es, WorkbookMethods nt, IgnoreSharedFormula i, SheetImpl si, WorkbookSettings ws):base(t, fr, si)
		{
			sbyte[] data = getRecord().Data;
			
			shared = false;
			
			// microsoft and their magic values determine whether this
			// is a string or a number value
			if (data[6] == 0 && data[12] == - 1 && data[13] == - 1)
			{
				// we have a string
				formula = new StringFormulaRecord(t, excelFile, fr, es, nt, si, ws);
			}
			else if (data[6] == 1 && data[12] == - 1 && data[13] == - 1)
			{
				// We have a boolean formula
				// multiple values.  Thanks to Frank for spotting this
				formula = new BooleanFormulaRecord(t, fr, es, nt, si);
			}
			else if (data[6] == 2 && data[12] == - 1 && data[13] == - 1)
			{
				// The cell is in error
				formula = new ErrorFormulaRecord(t, fr, es, nt, si);
			}
			else
			{
				// it is most assuredly a number
				formula = new NumberFormulaRecord(t, fr, es, nt, si);
			}
		}
		static FormulaRecord()
		{
			logger = Logger.getLogger(typeof(FormulaRecord));
		}
	}
}
