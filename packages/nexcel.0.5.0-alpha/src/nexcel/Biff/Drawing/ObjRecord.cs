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
using NExcel.Read.Biff;
namespace NExcel.Biff.Drawing
{
	
	/// <summary> A record which merely holds the OBJ data.  Used when copying files which
	/// contain images
	/// </summary>
	public class ObjRecord:WritableRecordData
	{
		/// <summary> Accessor for the object type
		/// 
		/// </summary>
		/// <returns> the object type
		/// </returns>
		virtual public ObjType Type
		{
			get
			{
				return type;
			}
			
		}
		/// <summary> The object type</summary>
		private ObjType type;
		
		/// <summary> Indicates whether this record was read in</summary>
		private bool read;
		
		/// <summary> The object id</summary>
		private int objectId;
		
		/// <summary> Object type enumeration</summary>
		public sealed class ObjType
		{
			internal int Value;
			internal ObjType(int v)
			{
				Value = v;
			}
		}
		
		/// <summary> A picture type indicator</summary>
		public static readonly ObjType PICTURE = new ObjType(0x08);
		
		/// <summary> A chart type indicator</summary>
		public static readonly ObjType CHART = new ObjType(0x05);
		
		// Field sub records
		private const int COMMON_DATA_LENGTH = 22;
		private const int CLIPBOARD_FORMAT_LENGTH = 6;
		private const int PICTURE_OPTION_LENGTH = 6;
		private const int END_LENGTH = 4;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public ObjRecord(Record t):base(t)
		{
			sbyte[] data = t.Data;
			int objtype = IntegerHelper.getInt(data[4], data[5]);
			read = true;
			
			if (objtype == CHART.Value)
			{
				type = CHART;
			}
			else if (objtype == PICTURE.Value)
			{
				type = PICTURE;
			}
			
			objectId = IntegerHelper.getInt(data[6], data[7]);
		}
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="objId">the object id
		/// </param>
		internal ObjRecord(int objId):base(NExcel.Biff.Type.OBJ)
		{
			objectId = objId;
			type = PICTURE;
		}
		
		/// <summary> Expose the protected function to the SheetImpl in this package
		/// 
		/// </summary>
		/// <returns> the raw record data
		/// </returns>
		public override sbyte[] getData()
		{
			if (read)
			{
				return getRecord().Data;
			}
			
			int dataLength = COMMON_DATA_LENGTH + CLIPBOARD_FORMAT_LENGTH + PICTURE_OPTION_LENGTH + END_LENGTH;
			int pos = 0;
			sbyte[] data = new sbyte[dataLength];
			
			// The common data
			// record id
			IntegerHelper.getTwoBytes(0x15, data, pos);
			
			// record .Length
			IntegerHelper.getTwoBytes(COMMON_DATA_LENGTH - 4, data, pos + 2);
			
			// object type
			IntegerHelper.getTwoBytes(type.Value, data, pos + 4);
			
			// object id
			IntegerHelper.getTwoBytes(objectId, data, pos + 6);
			
			// the options
			IntegerHelper.getTwoBytes(0x6011, data, pos + 8);
			pos += COMMON_DATA_LENGTH;
			
			// The clipboard format
			// record id
			IntegerHelper.getTwoBytes(0x7, data, pos);
			
			// record .Length
			IntegerHelper.getTwoBytes(CLIPBOARD_FORMAT_LENGTH - 4, data, pos + 2);
			
			// the data
			IntegerHelper.getTwoBytes(0xffff, data, pos + 4);
			pos += CLIPBOARD_FORMAT_LENGTH;
			
			// Picture option flags
			// record id
			IntegerHelper.getTwoBytes(0x8, data, pos);
			
			// record .Length
			IntegerHelper.getTwoBytes(PICTURE_OPTION_LENGTH - 4, data, pos + 2);
			
			// the data
			IntegerHelper.getTwoBytes(0x1, data, pos + 4);
			pos += CLIPBOARD_FORMAT_LENGTH;
			
			// End
			// record id
			IntegerHelper.getTwoBytes(0x0, data, pos);
			
			// record .Length
			IntegerHelper.getTwoBytes(END_LENGTH - 4, data, pos + 2);
			
			// the data
			pos += END_LENGTH;
			
			return data;
		}
		
		/// <summary> Expose the protected function to the SheetImpl in this package
		/// 
		/// </summary>
		/// <returns> the raw record data
		/// </returns>
		public override Record getRecord()
		{
			return base.getRecord();
		}
	}
}
