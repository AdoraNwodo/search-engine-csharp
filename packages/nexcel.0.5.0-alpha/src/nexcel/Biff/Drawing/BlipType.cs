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
namespace NExcel.Biff.Drawing
{
	
	/// <summary> Enumeration for the BLIP type</summary>
	sealed class BlipType
	{
		public string Description
		{
			get
			{
				return desc;
			}
			
		}
		private int _Value;
		private string desc;
		
		private static BlipType[] types;
		
		private BlipType(int val, string d)
		{
			_Value = val;
			desc = d;
			
			BlipType[] newtypes = new BlipType[types.Length + 1];
			Array.Copy(types, 0, newtypes, 0, types.Length);
			newtypes[types.Length] = this;
			types = newtypes;
		}
		
		public int Value
		{
			get
			{
				return _Value;
			}
		}
		
		public static BlipType getType(int val)
		{
			BlipType type = UNKNOWN;
			for (int i = 0; i < types.Length; i++)
			{
				if (types[i]._Value == val)
				{
					type = types[i];
					break;
				}
			}
			
			return type;
		}
		
		public static readonly BlipType ERROR = new BlipType(0, "Error"); // An error occured during loading
		public static readonly BlipType UNKNOWN = new BlipType(1, "Unknown"); // An unknown blip type
		public static readonly BlipType EMF = new BlipType(2, "EMF"); // Windows Enhanced Metafile
		public static readonly BlipType WMF = new BlipType(3, "WMF"); // Windows Metafile
		public static readonly BlipType PICT = new BlipType(4, "PICT"); // Macintosh PICT
		public static readonly BlipType JPEG = new BlipType(5, "JPEG"); // JFIF
		public static readonly BlipType PNG = new BlipType(6, "PNG"); // PNG
		public static readonly BlipType DIB = new BlipType(7, "DIB"); // Windows DIB
		public static readonly BlipType FIRST_CLIENT = new BlipType(32, "FIRST"); // First client defined blip type
		public static readonly BlipType LAST_CLIENT = new BlipType(255, "LAST"); // Last client defined blip type
		static BlipType()
		{
			types = new BlipType[0];
		}
	}
}
