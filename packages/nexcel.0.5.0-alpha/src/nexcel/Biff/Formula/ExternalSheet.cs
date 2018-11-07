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
using NExcel.Read.Biff;
namespace NExcel.Biff.Formula
{
	
	/// <summary> Interface which exposes the methods needed by formulas 
	/// to access external sheet records
	/// </summary>
	public interface ExternalSheet
		{
			/// <summary> Parsing of formulas is only supported for a subset of the available
			/// biff version, so we need to test to see if this version is acceptable
			/// 
			/// </summary>
			/// <returns> the BOF record, which 
			/// </returns>
			BOFRecord WorkbookBof
			{
				get;
				
			}
			/// <summary> Gets the name of the external sheet specified by the index
			/// 
			/// </summary>
			/// <param name="index">the external sheet index
			/// </param>
			/// <returns> the name of the external sheet
			/// </returns>
			string getExternalSheetName(int index);
			
			/// <summary> Gets the index of the first external sheet for the name
			/// 
			/// </summary>
			/// <param name="sheetName">the name of the external sheet
			/// </param>
			/// <returns>  the index of the external sheet with the specified name
			/// </returns>
			int getExternalSheetIndex(string sheetName);
			
			/// <summary> Gets the index of the first external sheet for the name
			/// 
			/// </summary>
			/// <param name="index">the external sheet index
			/// </param>
			/// <returns> the sheet index of the external sheet index
			/// </returns>
			int getExternalSheetIndex(int index);
			
			/// <summary> Gets the index of the last external sheet for the name
			/// 
			/// </summary>
			/// <param name="sheetName">the name of the external sheet
			/// </param>
			/// <returns>  the index of the external sheet with the specified name
			/// </returns>
			int getLastExternalSheetIndex(string sheetName);
			
			/// <summary> Gets the index of the first external sheet for the name
			/// 
			/// </summary>
			/// <param name="index">the external sheet index
			/// </param>
			/// <returns> the sheet index of the external sheet index
			/// </returns>
			int getLastExternalSheetIndex(int index);
		}
}
