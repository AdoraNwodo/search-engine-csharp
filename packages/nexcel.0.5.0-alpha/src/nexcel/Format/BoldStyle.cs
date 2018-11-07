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
	
	/// <summary> Enumeration class containing the various bold styles for data</summary>
	public class BoldStyle
	{
		/// <summary> Gets the value of the bold weight.  This is the value that will be
		/// written to the generated Excel file.
		/// 
		/// </summary>
		/// <returns> the bold weight
		/// </returns>
		virtual public int Value
		{
			get
			{
				return value_Renamed;
			}
			
		}
		/// <summary> Gets the string description of the bold style</summary>
		virtual public string Description
		{
			get
			{
				return string_Renamed;
			}
			
		}
		/// <summary> The bold weight</summary>
		private int value_Renamed;
		
		/// <summary> The description</summary>
		private string string_Renamed;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="val">
		/// </param>
		protected internal BoldStyle(int val, string s)
		{
			value_Renamed = val;
			string_Renamed = s;
		}
		
		/// <summary> Normal style</summary>
		public static readonly BoldStyle NORMAL = new BoldStyle(0x190, "Normal");
		/// <summary> Emboldened style</summary>
		public static readonly BoldStyle BOLD = new BoldStyle(0x2bc, "Bold");
	}
}