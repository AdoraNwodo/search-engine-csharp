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
namespace NExcel.Format
{
	
	/// <summary> Interface for cell formats</summary>
	public interface CellFormat
		{
			/// <summary> Gets the format used by this format
			/// 
			/// </summary>
			/// <returns> the format
			/// </returns>
			Format Format
			{
				get;
				
			}
			/// <summary> Gets the font information used by this format
			/// 
			/// </summary>
			/// <returns> the font
			/// </returns>
			Font Font
			{
				get;
				
			}
			/// <summary> Gets whether or not the contents of this cell are wrapped
			/// 
			/// </summary>
			/// <returns> TRUE if this cell's contents are wrapped, FALSE otherwise
			/// </returns>
			bool Wrap
			{
				get;
				
			}
			/// <summary> Gets the horizontal cell alignment
			/// 
			/// </summary>
			/// <returns> the alignment
			/// </returns>
			Alignment Alignment
			{
				get;
				
			}
			/// <summary> Gets the vertical cell alignment
			/// 
			/// </summary>
			/// <returns> the alignment
			/// </returns>
			VerticalAlignment VerticalAlignment
			{
				get;
				
			}
			/// <summary> Gets the orientation
			/// 
			/// </summary>
			/// <returns> the orientation
			/// </returns>
			Orientation Orientation
			{
				get;
				
			}
			/// <summary> Gets the background colour used by this cell
			/// 
			/// </summary>
			/// <returns> the foreground colour
			/// </returns>
			Colour BackgroundColour
			{
				get;
				
			}
			/// <summary> Gets the pattern used by this cell format
			/// 
			/// </summary>
			/// <returns> the background pattern
			/// </returns>
			Pattern Pattern
			{
				get;
				
			}
			/// <summary> Gets the shrink to fit flag
			/// 
			/// </summary>
			/// <returns> TRUE if this format is shrink to fit, FALSE otherise
			/// </returns>
			bool ShrinkToFit
			{
				get;
				
			}
			/// <summary> Accessor for whether a particular cell is locked
			/// 
			/// </summary>
			/// <returns> TRUE if this cell is locked, FALSE otherwise
			/// </returns>
			bool Locked
			{
				get;
				
			}
			
			/// <summary> Gets the line style for the given cell border
			/// If a border type of ALL or NONE is specified, then a line style of
			/// NONE is returned
			/// 
			/// </summary>
			/// <param name="border">the cell border we are interested in
			/// </param>
			/// <returns> the line style of the specified border
			/// </returns>
			BorderLineStyle getBorder(Border border);
			
			/// <summary> Gets the line style for the given cell border
			/// If a border type of ALL or NONE is specified, then a line style of
			/// NONE is returned
			/// 
			/// </summary>
			/// <param name="border">the cell border we are interested in
			/// </param>
			/// <returns> the line style of the specified border
			/// </returns>
			BorderLineStyle getBorderLine(Border border);
			
			/// <summary> Gets the colour for the given cell border
			/// If a border type of ALL or NONE is specified, then a line style of
			/// NONE is returned
			/// If the specified cell does not have an associated line style, then
			/// the colour the line would be is still returned
			/// 
			/// </summary>
			/// <param name="border">the cell border we are interested in
			/// </param>
			/// <returns> the line style of the specified border
			/// </returns>
			Colour getBorderColour(Border border);
			
			/// <summary> Determines if this cell format has any borders at all.  Used to
			/// set the new borders when merging a group of cells
			/// 
			/// </summary>
			/// <returns> TRUE if this cell has any borders, FALSE otherwise
			/// </returns>
			bool hasBorders();
		}
}