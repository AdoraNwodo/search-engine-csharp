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
namespace NExcel.Biff
{
	/// <summary> An interface containing some common workbook methods.  This so that
	/// objects which are re-used for both readable and writable workbooks
	/// can still make the same method calls on a workbook
	/// </summary>
	public interface WorkbookMethods
		{
			/// <summary> Gets the specified sheet within this workbook
			/// 
			/// </summary>
			/// <param name="index">the zero based index of the required sheet
			/// </param>
			/// <returns> The sheet specified by the index
			/// </returns>
			Sheet getReadSheet(int index);
			
			/// <summary> Gets the name at the specified index
			/// 
			/// </summary>
			/// <param name="index">the index into the name table
			/// </param>
			/// <returns> the name of the cell
			/// </returns>
			string getName(int index);
			
			/// <summary> Gets the index of the name record for the name
			/// 
			/// </summary>
			/// <param name="name">the name
			/// </param>
			/// <returns> the index in the name table
			/// </returns>
			int getNameIndex(string name);
		}
}
