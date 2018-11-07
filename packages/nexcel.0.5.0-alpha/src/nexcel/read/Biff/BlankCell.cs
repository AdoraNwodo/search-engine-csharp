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
namespace NExcel.Read.Biff
{
	
	/// <summary> A blank cell.  Despite the fact that this cell has no contents, it
	/// has formatting information applied to it
	/// </summary>
	public class BlankCell:CellValue
	{
		/// <summary> Returns the contents of this cell as an empty string
		/// 
		/// </summary>
		/// <returns> a empty string ""
		/// </returns>
		virtual public string Contents
		{
			get
			{
				return "";
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
				return CellType.EMPTY;
			}
			
		}
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="fr">the available formats
		/// </param>
		/// <param name="si">the sheet
		/// </param>
		internal BlankCell(Record t, FormattingRecords fr, SheetImpl si):base(t, fr, si)
		{
		}
	}
}
