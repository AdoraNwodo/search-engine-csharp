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
using System.Collections;
using common;
using NExcel;
using NExcel.Biff;
using NExcel.Biff.Formula;
namespace NExcel.Read.Biff
{
	
	/// <summary> A shared formula</summary>
	class SharedFormulaRecord
	{
		/// <summary> Accessor for the template formula.  Called when a shared formula has,
		/// for some reason, specified an inappropriate range and it is necessary
		/// to retrieve the template from a previously available shared formula
		/// </summary>
		virtual internal BaseSharedFormulaRecord TemplateFormula
		{
			get
			{
				return templateFormula;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The first row to which this shared formula applies</summary>
		private int firstRow;
		
		/// <summary> The last row to which this shared formula applies</summary>
		private int lastRow;
		
		/// <summary> The first column to which this shared formula applies</summary>
		private int firstCol;
		
		/// <summary> The last column to which this shared formula applies</summary>
		private int lastCol;
		
		/// <summary> The first (template) formula comprising this group</summary>
		private BaseSharedFormulaRecord templateFormula;
		
		/// <summary> The rest of the  comprising this shared formula</summary>
		private ArrayList formulas;
		
		/// <summary> The token data</summary>
		private sbyte[] tokens;
		
		/// <summary> A handle to the external sheet</summary>
		private ExternalSheet externalSheet;
		
		/// <summary> A handle to the name table</summary>
		private WorkbookMethods nameTable;
		
		/// <summary> A handle to the sheet</summary>
		private SheetImpl sheet;
		
		
		/// <summary> Constructs this object from the raw data.  Creates either a
		/// NumberFormulaRecord or a StringFormulaRecord depending on whether
		/// this formula represents a numerical calculation or not
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="fr">the base shared formula
		/// </param>
		/// <param name="es">the workbook, which contains the external sheet references
		/// </param>
		/// <param name="nt">the workbook
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		public SharedFormulaRecord(Record t, BaseSharedFormulaRecord fr, ExternalSheet es, WorkbookMethods nt, SheetImpl si)
		{
			externalSheet = es;
			nameTable = nt;
			sheet = si;
			sbyte[] data = t.Data;
			
			firstRow = IntegerHelper.getInt(data[0], data[1]);
			lastRow = IntegerHelper.getInt(data[2], data[3]);
			firstCol = (int) (data[4] & 0xff);
			lastCol = (int) (data[5] & 0xff);
			
			formulas = new ArrayList();
			
			templateFormula = fr;
			
			tokens = new sbyte[data.Length - 10];
			Array.Copy(data, 10, tokens, 0, tokens.Length);
		}
		
		/// <summary> Adds this formula to the list of formulas, if it falls within
		/// the bounds
		/// 
		/// </summary>
		/// <param name="fr">the formula record to test for membership of this group
		/// </param>
		/// <returns> TRUE if the formulas was added, FALSE otherwise
		/// </returns>
		public virtual bool add(BaseSharedFormulaRecord fr)
		{
			if (fr.Row >= firstRow && fr.Row <= lastRow && fr.Column >= firstCol && fr.Column <= lastCol)
			{
				formulas.Add(fr);
				return true;
			}
			
			return false;
		}
		
		/// <summary> Manufactures individual cell formulas out the whole shared formula
		/// debacle
		/// 
		/// </summary>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="nf">flag indicating whether this uses the 1904 date system
		/// </param>
		/// <returns> an array of formulas to be added to the sheet
		/// </returns>
		internal virtual Cell[] getFormulas(FormattingRecords fr, bool nf)
		{
			Cell[] sfs = new Cell[formulas.Count + 1];
			
			// This can happen if there are many identical formulas in the
			// sheet and excel has not sliced and diced them exclusively
			if (templateFormula == null)
			{
				logger.warn("Shared formula template formula is null");
				return new Cell[0];
			}
			
			templateFormula.setTokens(tokens);
			
			// See if the template formula evaluates to date
			if (templateFormula.Type == CellType.NUMBER_FORMULA)
			{
				if (fr.isDate(templateFormula.XFIndex))
				{
					SharedNumberFormulaRecord snfr = (SharedNumberFormulaRecord) templateFormula;
					templateFormula = new SharedDateFormulaRecord(snfr, fr, nf, sheet, snfr.FilePos);
					templateFormula.setTokens(snfr.getTokens());
				}
			}
			
			sfs[0] = templateFormula;
			
			BaseSharedFormulaRecord f = null;
			
			for (int i = 0; i < formulas.Count; i++)
			{
				f = (BaseSharedFormulaRecord) formulas[i];
				
				// See if the formula evaluates to date
				if (f.Type == CellType.NUMBER_FORMULA)
				{
					if (fr.isDate(f.XFIndex))
					{
						SharedNumberFormulaRecord snfr = (SharedNumberFormulaRecord) f;
						f = new SharedDateFormulaRecord(snfr, fr, nf, sheet, snfr.FilePos);
					}
				}
				
				f.setTokens(tokens);
				sfs[i + 1] = f;
			}
			
			return sfs;
		}
		static SharedFormulaRecord()
		{
			logger = Logger.getLogger(typeof(SharedFormulaRecord));
		}
	}
}
