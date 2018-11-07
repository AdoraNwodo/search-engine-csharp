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
	
	/// <summary> Enumeration class which contains the various alignments for data within a 
	/// cell
	/// </summary>
	public class Alignment
	{
		/// <summary> Gets the value of this alignment.  This is the value that is written to 
		/// the generated Excel file
		/// 
		/// </summary>
		/// <returns> the binary value
		/// </returns>
		virtual public int Value
		{
			get
			{
				return value_Renamed;
			}
			
		}
		/// <summary> Gets the string description of this alignment
		/// 
		/// </summary>
		/// <returns> the string description
		/// </returns>
		virtual public string Description
		{
			get
			{
				return string_Renamed;
			}
			
		}
		/// <summary> The internal numerical repreentation of the alignment</summary>
		private int value_Renamed;
		
		/// <summary> The string description of this alignment</summary>
		private string string_Renamed;
		
		/// <summary> The list of alignments</summary>
		private static Alignment[] alignments;
		
		/// <summary> Private constructor
		/// 
		/// </summary>
		/// <param name="val">
		/// </param>
		/// <param name="">string
		/// </param>
		protected internal Alignment(int val, string s)
		{
			value_Renamed = val;
			string_Renamed = s;
			
			Alignment[] oldaligns = alignments;
			alignments = new Alignment[oldaligns.Length + 1];
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
		public static Alignment getAlignment(int val)
		{
			for (int i = 0; i < alignments.Length; i++)
			{
				if (alignments[i].Value == val)
				{
					return alignments[i];
				}
			}
			
			return GENERAL;
		}
		
		/// <summary> The standard alignment</summary>
		public static Alignment GENERAL;
		/// <summary> Data cells with this alignment will appear at the left hand edge of the 
		/// cell
		/// </summary>
		public static Alignment LEFT;
		/// <summary> Data in cells with this alignment will be centred</summary>
		public static Alignment CENTRE;
		/// <summary> Data in cells with this alignment will be right aligned</summary>
		public static Alignment RIGHT;
		/// <summary> Data in cells with this alignment will fill the cell</summary>
		public static Alignment FILL;
		/// <summary> Data in cells with this alignment will be justified</summary>
		public static Alignment JUSTIFY;
		static Alignment()
		{
			alignments = new Alignment[0];
			GENERAL = new Alignment(0, "general");
			LEFT = new Alignment(1, "left");
			CENTRE = new Alignment(2, "centre");
			RIGHT = new Alignment(3, "right");
			FILL = new Alignment(4, "fill");
			JUSTIFY = new Alignment(5, "justify");
		}
	}
}