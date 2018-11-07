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
namespace NExcel.Read.Biff
{
	
	/// <summary> A container for the raw record data within a biff file</summary>
	public sealed class Record 
	{
		/// <summary> Gets the .Length of the record
		/// 
		/// </summary>
		/// <returns> the .Length of the record
		/// </returns>
		public int Length
		{
			get
			{
				return length;
			}
			
		}
		/// <summary> Gets the data portion of the record
		/// 
		/// </summary>
		/// <returns> the data portion of the record
		/// </returns>
		public sbyte[] Data
		{
			get
			{
				if (data == null)
				{
					data = file.read(dataPos, length);
				}
				
				return data;
			}
			
		}
		/// <summary> The excel 97 code
		/// 
		/// </summary>
		/// <returns> the excel code
		/// </returns>
		public int Code
		{
			get
			{
				return code;
			}
			
		}
		/// <summary> The excel biff code</summary>
		private int code;
		/// <summary> The data type</summary>
		private NExcel.Biff.Type type;
		/// <summary> The .Length of this record</summary>
		private int length;
		/// <summary> A pointer to the beginning of the actual data</summary>
		private int dataPos;
		/// <summary> A handle to the excel 97 file</summary>
		private File file;
		/// <summary> The raw data within this record</summary>
		private sbyte[] data;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="offset">the offset in the raw file
		/// </param>
		/// <param name="f">the excel 97 biff file
		/// </param>
		/// <param name="d">the data record
		/// </param>
		internal Record(sbyte[] d, int offset, File f)
		{
			code = IntegerHelper.getInt(d[offset], d[offset + 1]);
			length = IntegerHelper.getInt(d[offset + 2], d[offset + 3]);
			file = f;
			file.skip(4);
			dataPos = f.Pos;
			file.skip(length);
			type = NExcel.Biff.Type.getType(code);
		}
		
		/// <summary> Gets and sets the biff type
		/// 
		/// </summary>
		/// <returns> the biff type
		/// </returns>
		public NExcel.Biff.Type Type
		{
			get
			{
				return type;
			}
			set
			{
				type = value;
			}
		}
		
	}
}
