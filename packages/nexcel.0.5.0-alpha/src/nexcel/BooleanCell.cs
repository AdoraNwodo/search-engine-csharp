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
namespace NExcel
{
	
	/// <summary> This type represents the Microsoft concept of a Boolean.  Accordingly, this
	/// cell represents either TRUE, FALSE or an error condition.  This third
	/// state naturally makes handling BooleanCells quite tricky, and use of
	/// the specific access methods should be handled with care
	/// </summary>
	public interface BooleanCell : Cell
		{
			/// <summary> Gets the boolean value stored in this cell.  If this cell contains an
			/// error, then returns FALSE.  Always query this cell type using the
			/// accessor method isError() prior to calling this method
			/// 
			/// </summary>
			/// <returns> TRUE if this cell contains TRUE, FALSE if it contains FALSE or
			/// an error code
			/// </returns>
			bool BooleanValue
			{
				get;
			}
		}
}
