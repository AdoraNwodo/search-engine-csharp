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
	
	/// <summary> Interface which exposes the user font display information to the user</summary>
	public interface Font
		{
			/// <summary> Gets the name of this font
			/// 
			/// </summary>
			/// <returns> the name of this font
			/// </returns>
			string Name
			{
				get;
				
			}
			/// <summary> Gets the point size for this font, if the font hasn't been initialized
			/// 
			/// </summary>
			/// <returns> the point size
			/// </returns>
			int PointSize
			{
				get;
				
			}
			/// <summary> Gets the bold weight for this font
			/// 
			/// </summary>
			/// <returns> the bold weight for this font
			/// </returns>
			int BoldWeight
			{
				get;
				
			}
			/// <summary> Returns the italic flag
			/// 
			/// </summary>
			/// <returns> TRUE if this font is italic, FALSE otherwise
			/// </returns>
			bool Italic
			{
				get;
				
			}
			/// <summary> Gets the underline style for this font
			/// 
			/// </summary>
			/// <returns> the underline style
			/// </returns>
			UnderlineStyle UnderlineStyle
			{
				get;
				
			}
			/// <summary> Gets the colour for this font
			/// 
			/// </summary>
			/// <returns> the colour
			/// </returns>
			Colour Colour
			{
				get;
				
			}
			/// <summary> Gets the script style
			/// 
			/// </summary>
			/// <returns> the script style
			/// </returns>
			ScriptStyle ScriptStyle
			{
				get;
				
			}
		}
}