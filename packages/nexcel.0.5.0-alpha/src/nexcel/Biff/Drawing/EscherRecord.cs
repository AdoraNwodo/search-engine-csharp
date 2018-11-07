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
	
	/// <summary> The base class for all escher records.  This class contains
	/// the common header data and is basically a wrapper for the EscherRecordData
	/// object
	/// </summary>
	public abstract class EscherRecord
	{
		/// <summary> Identifies whether this item is a container
		/// 
		/// </summary>
		/// <param name="cont">TRUE if this is a container, FALSE otherwise
		/// </param>
		virtual protected internal bool Container
		{
			set
			{
				data.setContainer(value);
			}
			
		}
		/// <summary> Gets the entire .Length of the record, including the header
		/// 
		/// </summary>
		/// <returns> the .Length of the record, including the header data
		/// </returns>
		virtual public int Length
		{
			get
			{
				return data.getLength() + HEADER_LENGTH;
			}
			
		}
		/// <summary> Accessor for the escher stream 
		/// 
		/// </summary>
		/// <returns> the escher stream
		/// </returns>
		virtual protected internal EscherStream EscherStream
		{
			get
			{
				return data.EscherStream;
			}
			
		}
		/// <summary> The position of this escher record in the stream
		/// 
		/// </summary>
		/// <returns> the position
		/// </returns>
		virtual protected internal int Pos
		{
			get
			{
				return data.Pos;
			}
			
		}
		/// <summary> Accessor for the instance
		/// 
		/// </summary>
		/// <returns> the instance
		/// </returns>
		/// <summary> Sets the instance number when writing out the escher data
		/// 
		/// </summary>
		/// <param name="i">the instance
		/// </param>
		virtual protected internal int Instance
		{
			get
			{
				return data.Instance;
			}
			
			set
			{
				data.Instance = value;
			}
			
		}
		/// <summary> Sets the version when writing out the escher data
		/// 
		/// </summary>
		/// <param name="v">the version
		/// </param>
		virtual protected internal int Version
		{
			set
			{
				data.Version = value;
			}
			
		}
		/// <summary> Accessor for the escher type
		/// 
		/// </summary>
		/// <returns> the type
		/// </returns>
		virtual public EscherRecordType Type
		{
			get
			{
				return data.Type;
			}
			
		}
		/// <summary> Abstract method used to retrieve the generated escher data when writing
		/// out image information
		/// 
		/// </summary>
		/// <returns> the escher data
		/// </returns>
		public abstract sbyte[] Data{get;}
		/// <summary> Gets the data that was read in, excluding the header data
		/// 
		/// </summary>
		/// <returns> the bytes read in, excluding the header data
		/// </returns>
		virtual internal sbyte[] Bytes
		{
			get
			{
				return data.Bytes;
			}
			
		}
		/// <summary> The escher data</summary>
		private EscherRecordData data;
		
		/// <summary> The .Length of the escher header on all records</summary>
		protected internal const int HEADER_LENGTH = 8;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="erd">the data
		/// </param>
		protected internal EscherRecord(EscherRecordData erd)
		{
			data = erd;
		}
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="type">the type
		/// </param>
		protected internal EscherRecord(EscherRecordType type)
		{
			data = new EscherRecordData(type);
		}
		
		/// <summary> Prepends the standard header data to the first eight bytes of the array
		/// and returns it
		/// </summary>
		internal sbyte[] setHeaderData(sbyte[] d)
		{
			return data.setHeaderData(d);
		}
	}
}
