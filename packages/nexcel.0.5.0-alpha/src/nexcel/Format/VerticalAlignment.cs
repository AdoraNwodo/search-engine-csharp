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
	
	/// <summary> Enumeration type which describes the vertical alignment of data within a cell</summary>
	public class VerticalAlignment
	{
		/// <summary> Accessor for the binary value
		/// 
		/// </summary>
		/// <returns> the internal binary value
		/// </returns>
		virtual public int Value
		{
			get
			{
				return value_Renamed;
			}
			
		}
		/// <summary> Gets the textual description</summary>
		virtual public string Description
		{
			get
			{
				return string_Renamed;
			}
			
		}
		/// <summary> The internal binary value which gets written to the generated Excel file</summary>
		private int value_Renamed;
		
		/// <summary> The textual description</summary>
		private string string_Renamed;
		
		/// <summary> The list of alignments</summary>
		private static VerticalAlignment[] alignments;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="val">
		/// </param>
		protected internal VerticalAlignment(int val, string s)
		{
			value_Renamed = val;
			string_Renamed = s;
			
			VerticalAlignment[] oldaligns = alignments;
			alignments = new VerticalAlignment[oldaligns.Length + 1];
			Array.Copy((System.Array) oldaligns, 0, (System.Array) alignments, 0, oldaligns.Length);
			alignments[oldaligns.Length] = this;
		}
		
		/// <summary> Gets the alignment from the value
		/// 
		/// </summary>
		/// <param name="val">
		/// </param>
		/// <returns> the alignment with that value
		/// </returns>
		public static VerticalAlignment getAlignment(int val)
		{
			for (int i = 0; i < alignments.Length; i++)
			{
				if (alignments[i].Value == val)
				{
					return alignments[i];
				}
			}
			
			return BOTTOM;
		}
		
		
		/// <summary> Cells with this specified vertical alignment will have their data
		/// aligned at the top
		/// </summary>
		public static VerticalAlignment TOP;
		/// <summary> Cells with this specified vertical alignment will have their data
		/// aligned centrally
		/// </summary>
		public static VerticalAlignment CENTRE;
		/// <summary> Cells with this specified vertical alignment will have their data
		/// aligned at the bottom
		/// </summary>
		public static VerticalAlignment BOTTOM;
		/// <summary> Cells with this specified vertical alignment will have their data
		/// justified
		/// </summary>
		public static VerticalAlignment JUSTIFY;
		static VerticalAlignment()
		{
			alignments = new VerticalAlignment[0];
			TOP = new VerticalAlignment(0, "top");
			CENTRE = new VerticalAlignment(1, "centre");
			BOTTOM = new VerticalAlignment(2, "bottom");
			JUSTIFY = new VerticalAlignment(3, "Justify");
		}
	}
}