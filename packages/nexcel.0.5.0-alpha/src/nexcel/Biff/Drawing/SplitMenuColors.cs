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
	
	class SplitMenuColors:EscherAtom
	{
		new private sbyte[] data;
		
		public SplitMenuColors(EscherRecordData erd):base(erd)
		{
		}
		
		public SplitMenuColors():base(EscherRecordType.SPLIT_MENU_COLORS)
		{
			Version = 0;
			Instance = 4;
			
			data = new sbyte[]{
								  (sbyte) 0x0d, 
								  (sbyte) 0x00, 
								  (sbyte) 0x00, 
								  (sbyte) 0x08, 
								  (sbyte) 0x0c, 
								  (sbyte) 0x00, 
								  (sbyte) 0x00, 
								  (sbyte) 0x08, 
								  (sbyte) 0x17, 
								  (sbyte) 0x00, 
								  (sbyte) 0x00, 
								  (sbyte) 0x08, 
								  (sbyte) -0x09, 
								  (sbyte) 0x00, 
								  (sbyte) 0x00, 
								  (sbyte) 0x10
							  };
		}
		
		public override sbyte[] Data
		{
		get
		{
		return setHeaderData(data);
		}
		}
	}
}
