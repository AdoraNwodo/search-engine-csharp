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
using System.Text;
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	/// <summary> A boundsheet record, which contains the worksheet name</summary>
	class BoundsheetRecord:RecordData
	{
		/// <summary> Accessor for the worksheet name
		/// 
		/// </summary>
		/// <returns> the worksheet name
		/// </returns>
		virtual public string Name
		{
			get
			{
				return name;
			}
			
		}
		/// <summary> Accessor for the hidden flag
		/// 
		/// </summary>
		/// <returns> TRUE if this is a hidden sheet, FALSE otherwise
		/// </returns>
		virtual public bool isHidden()
		{
				return visibilityFlag != 0;
		}

		/// <summary> Accessor to determine if this is a worksheet, or some other nefarious
		/// type of object
		/// 
		/// </summary>
		/// <returns> TRUE if this is a worksheet, FALSE otherwise
		/// </returns>
		virtual public bool isSheet()
		{
				return typeFlag == 0;
		}

		/// <summary> Accessor to determine if this is a chart
		/// 
		/// </summary>
		/// <returns> TRUE if this is a chart, FALSE otherwise
		/// </returns>
		virtual public bool Chart
		{
			get
			{
				return typeFlag == 2;
			}
			
		}
		/// <summary> The offset into the sheet</summary>
		private int offset;
		/// <summary> The type of sheet this is</summary>
		private sbyte typeFlag;
		/// <summary> The visibility flag</summary>
		private sbyte visibilityFlag;
		/// <summary> The length of the worksheet name</summary>
		private int length;
		/// <summary> The worksheet name</summary>
		private string name;
		
		/// <summary> Dummy indicators for overloading the constructor</summary>
		public class Biff7
		{
		}
		
		public static Biff7 biff7;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		public BoundsheetRecord(Record t):base(t)
		{
			sbyte[] data = getRecord().Data;
			offset = IntegerHelper.getInt(data[0], data[1], data[2], data[3]);
			typeFlag = data[5];
			visibilityFlag = data[4];
			length = data[6];
			
			if (data[7] == 0)
			{
				// Standard ASCII encoding
				sbyte[] bytes = new sbyte[length];
				Array.Copy(data, 8, bytes, 0, length);
				name = new string(NExcelUtils.Byte.ToCharArray(NExcelUtils.Byte.ToByteArray(bytes)));
			}
			else
			{
				// little endian Unicode encoding
				sbyte[] bytes = new sbyte[length * 2];
				Array.Copy(data, 8, bytes, 0, length * 2);
				try
				{
					// [TODO] - test if it is right - critical
					// name = new String(bytes, "UnicodeLittle");
//					name = new string(NExcelUtils.Byte.ToCharArray(NExcelUtils.Byte.ToByteArray(bytes)));
					byte[] bb = NExcelUtils.Byte.ToByteArray(bytes);
					Encoding encoding = Encoding.Unicode;
					name = encoding.GetString(bb);
				}
				catch (System.Exception ex)
				{
					// fail silently
					name = "Error";
				}
			}
		}
		
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="biff7">a dummy value to tell the record to interpret the
		/// data as biff7
		/// </param>
		public BoundsheetRecord(Record t, Biff7 biff7):base(t)
		{
			sbyte[] data = getRecord().Data;
			offset = IntegerHelper.getInt(data[0], data[1], data[2], data[3]);
			typeFlag = data[5];
			visibilityFlag = data[4];
			length = data[6];
			sbyte[] bytes = new sbyte[length];
			Array.Copy(data, 7, bytes, 0, length);
			name = new string(NExcelUtils.Byte.ToCharArray(NExcelUtils.Byte.ToByteArray(bytes)));
		}
		static BoundsheetRecord()
		{
			biff7 = new Biff7();
		}
	}
}
