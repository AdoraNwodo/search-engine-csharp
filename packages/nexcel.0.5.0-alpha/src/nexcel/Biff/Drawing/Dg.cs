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
	
	/// <summary> The Drawing Group</summary>
	class Dg:EscherAtom
	{
		/// <summary> Gets the drawing id
		/// 
		/// </summary>
		/// <returns> the drawing id
		/// </returns>
		virtual public int DrawingId
		{
			get
			{
				return drawingId;
			}
			
		}
		/// <summary> The data</summary>
		new private sbyte[] data;
		
		/// <summary> The id of this drawing</summary>
		private int drawingId;
		
		/// <summary> The number of shapes</summary>
		private int shapeCount;
		
		/// <summary> The seed for drawing ids</summary>
		private int seed;
		
		/// <summary> Constructor invoked when reading in an escher stream
		/// 
		/// </summary>
		/// <param name="erd">the escher record
		/// </param>
		public Dg(EscherRecordData erd):base(erd)
		{
			drawingId = Instance;
			
			sbyte[] bytes = Bytes;
			shapeCount = IntegerHelper.getInt(bytes[0], bytes[1], bytes[2], bytes[3]);
			seed = IntegerHelper.getInt(bytes[4], bytes[5], bytes[6], bytes[7]);
		}
		
		/// <summary> Constructor invoked when writing out an escher stream
		/// 
		/// </summary>
		/// <param name="numDrawings">the number of drawings
		/// </param>
		public Dg(int numDrawings):base(EscherRecordType.DG)
		{
			drawingId = 1;
			shapeCount = numDrawings + 1;
			seed = 1024 + shapeCount + 1;
			Instance = drawingId;
		}
		
		/// <summary> Used to generate the drawing data
		/// 
		/// </summary>
		/// <returns> the data
		/// </returns>
		public override sbyte[] Data
		{
		get
		{
		data = new sbyte[8];
		IntegerHelper.getFourBytes(shapeCount, data, 0);
		IntegerHelper.getFourBytes(seed, data, 4);
		
		return setHeaderData(data);
		}
		}
	}
}
