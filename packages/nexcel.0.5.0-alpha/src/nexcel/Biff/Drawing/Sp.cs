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
namespace NExcel.Biff.Drawing
{
	
	class Sp:EscherAtom
	{
		virtual internal int ShapeId
		{
			get
			{
				return shapeId;
			}
			
		}
		new private sbyte[] data;
		private int shapeType;
		private int shapeId;
		private int persistenceFlags;
		
		public class ShapeType
		{
			internal int Value;
			internal ShapeType(int v)
			{
				Value = v;
			}
		}
		public static readonly ShapeType MIN = new ShapeType(0);
		public static readonly ShapeType PICTURE_FRAME = new ShapeType(75);
		
		public Sp(EscherRecordData erd):base(erd)
		{
			shapeType = Instance;
			sbyte[] bytes = Bytes;
			shapeId = IntegerHelper.getInt(bytes[0], bytes[1], bytes[2], bytes[3]);
			persistenceFlags = IntegerHelper.getInt(bytes[4], bytes[5], bytes[6], bytes[7]);
		}
		
		public Sp(ShapeType st, int sid, int p):base(EscherRecordType.SP)
		{
			Version = 2;
			shapeType = st.Value;
			shapeId = sid;
			persistenceFlags = p;
			Instance = shapeType;
		}
		
		public override sbyte[] Data
		{
		get
		{
		data = new sbyte[8];
		IntegerHelper.getFourBytes(shapeId, data, 0);
		IntegerHelper.getFourBytes(persistenceFlags, data, 4);
		return setHeaderData(data);
		}
		}
	}
}
