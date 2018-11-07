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
	
	/// <summary> Enumeration class for Escher record types</summary>
	public sealed class EscherRecordType
	{
		/// <summary> The code of the item within the escher stream</summary>
		private int _Value;
		
		/// <summary> All escher types</summary>
		private static EscherRecordType[] types;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="val">the escher record value
		/// </param>
		private EscherRecordType(int val)
		{
			_Value = val;
			
			EscherRecordType[] newtypes = new EscherRecordType[types.Length + 1];
			Array.Copy(types, 0, newtypes, 0, types.Length);
			newtypes[types.Length] = this;
			types = newtypes;
		}
		
		/// <summary> Accessor for the escher record value
		/// 
		/// </summary>
		/// <returns> the escher record value
		/// </returns>
		public int Value
		{
			get
			{
				return _Value;
			}
		}
		
		/// <summary> Accessor to get the item from a particular value
		/// 
		/// </summary>
		/// <param name="val">the escher record value
		/// </param>
		/// <returns> the type corresponding to val, or UNKNOWN if a match could not
		/// be found
		/// </returns>
		public static EscherRecordType getType(int val)
		{
			EscherRecordType type = UNKNOWN;
			
			for (int i = 0; i < types.Length; i++)
			{
				if (val == types[i]._Value)
				{
					type = types[i];
					break;
				}
			}
			
			return type;
		}

		public static readonly EscherRecordType UNKNOWN;
		public static readonly EscherRecordType DGG_CONTAINER;
		public static readonly EscherRecordType BSTORE_CONTAINER;
		public static readonly EscherRecordType DG_CONTAINER;
		public static readonly EscherRecordType SPGR_CONTAINER;
		public static readonly EscherRecordType SP_CONTAINER;

		public static readonly EscherRecordType DGG;
		public static readonly EscherRecordType BSE;
		public static readonly EscherRecordType DG;
		public static readonly EscherRecordType SPGR;
		public static readonly EscherRecordType SP;
		public static readonly EscherRecordType OPT;
		public static readonly EscherRecordType CLIENT_ANCHOR;

		public static readonly EscherRecordType CLIENT_DATA;
		public static readonly EscherRecordType SPLIT_MENU_COLORS;


		static EscherRecordType()
		{
			types = new EscherRecordType[0];

			// Init values
			UNKNOWN = new EscherRecordType(0x0);
			DGG_CONTAINER = new EscherRecordType(0xf000);
			BSTORE_CONTAINER = new EscherRecordType(0xf001);
			DG_CONTAINER = new EscherRecordType(0xf002);
			SPGR_CONTAINER = new EscherRecordType(0xf003);
			SP_CONTAINER = new EscherRecordType(0xf004);

			DGG = new EscherRecordType(0xf006);
			BSE = new EscherRecordType(0xf007);
			DG = new EscherRecordType(0xf008);
			SPGR = new EscherRecordType(0xf009);
			SP = new EscherRecordType(0xf00a);
			OPT = new EscherRecordType(0xf00b);
			CLIENT_ANCHOR = new EscherRecordType(0xf010);

			CLIENT_DATA = new EscherRecordType(0xf011);
			SPLIT_MENU_COLORS = new EscherRecordType(0xf11e);
		}

	}
}
