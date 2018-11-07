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
using Uri = System.Uri;
namespace NExcel
{
	
	/// <summary> Hyperlink information.  Only URLs or file links are supported
	/// 
	/// Hyperlinks may apply to a range of cells; in such cases the methods
	/// getRow and getColumn return the cell at the top left of the range
	/// the hyperlink refers to.  Hyperlinks have no specific cell format
	/// information applied to them, so the getCellFormat method will return null
	/// </summary>
	public interface Hyperlink
		{
			/// <summary> Returns the row number of this cell
			/// 
			/// </summary>
			/// <returns> the row number of this cell
			/// </returns>
			int Row
			{
				get;
				
			}
			/// <summary> Returns the column number of this cell
			/// 
			/// </summary>
			/// <returns> the column number of this cell
			/// </returns>
			int Column
			{
				get;
				
			}
			/// <summary> Gets the range of cells which activate this hyperlink
			/// The get sheet index methods will all return -1, because the
			/// cells will all be present on the same sheet
			/// 
			/// </summary>
			/// <returns> the range of cells which activate the hyperlink
			/// </returns>
			Range Range
			{
				get;
				
			}
			/// <summary> Determines whether this is a hyperlink to a location in this workbook
			/// 
			/// </summary>
			/// <returns> TRUE if this is a link to an internal location
			/// </returns>
			bool isLocation();

			/// <summary> Returns the row number of the bottom right cell
			/// 
			/// </summary>
			/// <returns> the row number of this cell
			/// </returns>
			int LastRow
			{
				get;
				
			}
			/// <summary> Returns the column number of the bottom right cell
			/// 
			/// </summary>
			/// <returns> the column number of this cell
			/// </returns>
			int LastColumn
			{
				get;
				
			}
			
			/// <summary> Determines whether this is a hyperlink to a file
			/// 
			/// </summary>
			/// <returns> TRUE if this is a hyperlink to a file, FALSE otherwise
			/// </returns>
			bool isFile();
			
			/// <summary> Determines whether this is a hyperlink to a web resource
			/// 
			/// </summary>
			/// <returns> TRUE if this is a URL
			/// </returns>
			bool isURL();
			
			/// <summary> Gets the URL referenced by this Hyperlink
			/// 
			/// </summary>
			/// <returns> the URL, or NULL if this hyperlink is not a URL
			/// </returns>
			Uri getURL();
			
			/// <summary> Returns the local file eferenced by this Hyperlink
			/// 
			/// </summary>
			/// <returns> the file, or NULL if this hyperlink is not a file
			/// </returns>
			System.IO.FileInfo getFile();
		}
}
