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
namespace NExcel.Read.Biff
{
	
	/// <summary> Contains the cell dimensions of this worksheet</summary>
	class Window2Record:RecordData
	{
		/// <summary> Accessor for the selected flag
		/// 
		/// </summary>
		/// <returns> TRUE if this sheet is selected, FALSE otherwise
		/// </returns>
		virtual public bool Selected
		{
			get
			{
				return selected;
			}
			
		}
		/// <summary> Accessor for the show grid lines flag
		/// 
		/// </summary>
		/// <returns> TRUE to show grid lines, FALSE otherwise
		/// </returns>
		virtual public bool ShowGridLines
		{
			get
			{
				return showGridLines;
			}
			
		}
		/// <summary> Accessor for the zero values flag
		/// 
		/// </summary>
		/// <returns> TRUE if this sheet displays zero values, FALSE otherwise
		/// </returns>
		virtual public bool DisplayZeroValues
		{
			get
			{
				return displayZeroValues;
			}
			
		}
		/// <summary> Accessor for the frozen panes flag
		/// 
		/// </summary>
		/// <returns> TRUE if this contains frozen panes, FALSE otherwise
		/// </returns>
		virtual public bool Frozen
		{
			get
			{
				return frozenPanes;
			}
			
		}
		/// <summary> Accessor for the frozen not split flag
		/// 
		/// </summary>
		/// <returns> TRUE if this contains frozen, FALSE otherwise
		/// </returns>
		virtual public bool FrozenNotSplit
		{
			get
			{
				return frozenNotSplit;
			}
			
		}
		/// <summary> Selected flag</summary>
		private bool selected;
		/// <summary> Show grid lines flag</summary>
		private bool showGridLines;
		/// <summary> Display zero values flag</summary>
		private bool displayZeroValues;
		/// <summary> The window contains frozen panes</summary>
		private bool frozenPanes;
		/// <summary> The window contains panes that are frozen but not split</summary>
		private bool frozenNotSplit;
		
		/// <summary> Constructs the dimensions from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public Window2Record(Record t):base(t)
		{
			sbyte[] data = t.Data;
			
			sbyte options = data[0];
			sbyte sel = data[1];
			
			selected = ((sel & 0x02) != 0);
			showGridLines = ((options & 0x02) != 0);
			displayZeroValues = ((options & 0x10) != 0);
			frozenPanes = ((options & 0x08) != 0);
			frozenNotSplit = ((sel & 0x01) != 0);
		}
	}
}
