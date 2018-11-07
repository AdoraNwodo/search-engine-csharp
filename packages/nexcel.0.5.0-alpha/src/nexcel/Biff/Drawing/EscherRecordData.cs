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
	
	/// <summary> A single record from an Escher stream.  Basically this a container for
	/// the header data for each Escher record
	/// </summary>
	public sealed class EscherRecordData
	{
		/// <summary> Accessor for the record id
		/// 
		/// </summary>
		/// <returns> the record id
		/// </returns>
		public int RecordId
		{
			get
			{
				return recordId;
			}
			
		}
		/// <summary> Accessor for the drawing group stream
		/// 
		/// </summary>
		/// <returns> the drawing group stream
		/// </returns>
		internal EscherStream DrawingGroup
		{
			get
			{
				return escherStream;
			}
			
		}
		/// <summary> Gets the position in the stream
		/// 
		/// </summary>
		/// <returns> the position in the stream
		/// </returns>
		internal int Pos
		{
			get
			{
				return pos;
			}
			
		}
		/// <summary> Gets the escher type of this record</summary>
		internal EscherRecordType Type
		{
			get
			{
				if (type == null)
				{
					type = EscherRecordType.getType(recordId);
				}
				
				return type;
			}
			
		}
		/// <summary> Gets the instance value
		/// 
		/// </summary>
		/// <returns> the instance value
		/// </returns>
		/// <summary> Called from the subclass when writing to set the instance value
		/// 
		/// </summary>
		/// <param name="inst">the instance
		/// </param>
		internal int Instance
		{
			get
			{
				return instance;
			}
			
			set
			{
				instance = value;
			}
			
		}
		/// <summary> Called when writing to set the version of this record
		/// 
		/// </summary>
		/// <param name="v">the version
		/// </param>
		internal int Version
		{
			set
			{
				version = value;
			}
			
		}
		/// <summary> Accessor for the header stream 
		/// 
		/// </summary>
		/// <returns> the escher stream
		/// </returns>
		internal EscherStream EscherStream
		{
			get
			{
				return escherStream;
			}
			
		}
		/// <summary> Gets the data that was read in, excluding the header data
		/// 
		/// </summary>
		/// <returns> the value data that was read in
		/// </returns>
		internal sbyte[] Bytes
		{
			get
			{
				sbyte[] d = new sbyte[length];
				Array.Copy(escherStream.getData(), pos + 8, d, 0, length);
				return d;
			}
			
		}
		/// <summary> The byte position of this record in the escher stream</summary>
		private int pos;
		
		/// <summary> The instance value</summary>
		private int instance;
		
		/// <summary> The version value</summary>
		private int version;
		
		/// <summary> The record id</summary>
		private int recordId;
		
		/// <summary> The .Length of the record, excluding the 8 byte header</summary>
		private int length;
		
		/// <summary> Indicates whether this record is a container</summary>
		private bool container;
		
		/// <summary> The type of this record</summary>
		private EscherRecordType type;
		
		/// <summary> A handle back to the drawing group, which contains the entire escher
		/// stream byte data
		/// </summary>
		private EscherStream escherStream;
		
		/// <summary> Constructor</summary>
		public EscherRecordData(EscherStream dg, int p)
		{
			escherStream = dg;
			pos = p;
			sbyte[] data = escherStream.getData();
			
			// First two bytes contain instance and version
			int Value = IntegerHelper.getInt(data[pos], data[pos + 1]);
			
			// Instance value is the first 12 bits
			instance = (Value & 0xfff0) >> 4;
			
			// Version is the last four bits
			version = Value & 0xf;
			
			// Bytes 2 and 3 are the record id
			recordId = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
			
			// Length is bytes 4,5,6 and 7
			length = IntegerHelper.getInt(data[pos + 4], data[pos + 5], data[pos + 6], data[pos + 7]);
			
			if (version == 0x0f)
			{
				container = true;
			}
			else
			{
				container = false;
			}
		}
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="t">the type of the escher record
		/// </param>
		public EscherRecordData(EscherRecordType t)
		{
			type = t;
			recordId = type.Value;
		}
		
		/// <summary> Determines whether this record is a container
		/// 
		/// </summary>
		/// <returns> TRUE if this is a container, FALSE otherwise
		/// </returns>
		public bool isContainer()
		{
			return container;
		}
		
		/// <summary> Accessor for the .Length, excluding the 8 byte header
		/// 
		/// </summary>
		/// <returns> the .Length excluding the 8 byte header
		/// </returns>
		public int getLength()
		{
			return length;
		}
		
		/// <summary> Sets whether or not this is a container - called when writing
		/// out an escher stream
		/// 
		/// </summary>
		/// <param name="c">TRUE if this is a container, FALSE otherwise
		/// </param>
		internal void  setContainer(bool c)
		{
			container = c;
		}
		
		/// <summary> Called when writing to set the .Length of this record
		/// 
		/// </summary>
		/// <param name="l">the .Length
		/// </param>
		internal void  setLength(int l)
		{
			length = l;
		}
		
		/// <summary> Adds the 8 byte header data on the value data passed in, returning
		/// the modified data
		/// 
		/// </summary>
		/// <param name="d">the value data
		/// </param>
		/// <returns> the value data with the header information
		/// </returns>
		internal sbyte[] setHeaderData(sbyte[] d)
		{
			sbyte[] data = new sbyte[d.Length + 8];
			Array.Copy(d, 0, data, 8, d.Length);
			
			if (container)
			{
				version = 0x0f;
			}
			
			// First two bytes contain instance and version
			int Value = instance << 4;
			Value |= version;
			IntegerHelper.getTwoBytes(Value, data, 0);
			
			// Bytes 2 and 3 are the record id
			IntegerHelper.getTwoBytes(recordId, data, 2);
			
			// Length is bytes 4,5,6 and 7
			IntegerHelper.getFourBytes(d.Length, data, 4);
			
			return data;
		}
	}
}
