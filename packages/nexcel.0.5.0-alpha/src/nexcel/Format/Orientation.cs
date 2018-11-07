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
	
	/// <summary> Enumeration type which describes the orientation of data within a cell</summary>
	public sealed class Orientation
	{
		/// <summary> Accessor for the binary value
		/// 
		/// </summary>
		/// <returns> the internal binary value
		/// </returns>
		public int Value
		{
			get
			{
				return value_Renamed;
			}
			
		}
		/// <summary> Gets the textual description</summary>
		public string Description
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
		private static Orientation[] orientations;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="val">
		/// </param>
		protected internal Orientation(int val, string s)
		{
			value_Renamed = val; string_Renamed = s;
			
			Orientation[] oldorients = orientations;
			orientations = new Orientation[oldorients.Length + 1];
			Array.Copy((System.Array) oldorients, 0, (System.Array) orientations, 0, oldorients.Length);
			orientations[oldorients.Length] = this;
		}
		
		/// <summary> Gets the alignment from the value
		/// 
		/// </summary>
		/// <param name="val">
		/// </param>
		/// <returns> the alignment with that value
		/// </returns>
		public static Orientation getOrientation(int val)
		{
			for (int i = 0; i < orientations.Length; i++)
			{
				if (orientations[i].Value == val)
				{
					return orientations[i];
				}
			}
			
			return HORIZONTAL;
		}
		
		
		/// <summary> Cells with this specified orientation will be horizontal</summary>
		public static Orientation HORIZONTAL;
		/// <summary> Cells with this specified orientation have their data
		/// presented vertically
		/// </summary>
		public static Orientation VERTICAL;
		/// <summary> Cells with this specified orientation will have their data
		/// presented with a rotation of 90 degrees upwards
		/// </summary>
		public static Orientation PLUS_90;
		/// <summary> Cells with this specified orientation will have their data
		/// presented with a rotation of 90 degrees downwardswards
		/// </summary>
		public static Orientation MINUS_90;
		static Orientation()
		{
			orientations = new Orientation[0];
			HORIZONTAL = new Orientation(0, "horizontal");
			VERTICAL = new Orientation(0xff, "vertical");
			PLUS_90 = new Orientation(90, "up 90");
			MINUS_90 = new Orientation(180, "down 90");
		}
	}
}