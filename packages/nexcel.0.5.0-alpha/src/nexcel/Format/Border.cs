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
	
	/// <summary> The location of a border</summary>
	public class Border
	{
		/// <summary> Gets the description</summary>
		virtual public string Description
		{
			get
			{
				return string_Renamed;
			}
			
		}
		/// <summary> The string description</summary>
		private string string_Renamed;
		
		/// <summary> Constructor</summary>
		protected internal Border(string s)
		{
			string_Renamed = s;
		}
		
		public static readonly Border NONE = new Border("none");
		public static readonly Border ALL = new Border("all");
		public static readonly Border TOP = new Border("top");
		public static readonly Border BOTTOM = new Border("bottom");
		public static readonly Border LEFT = new Border("left");
		public static readonly Border RIGHT = new Border("right");
	}
}