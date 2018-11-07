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
using common;
using NExcel.Biff;
namespace NExcel.Biff.Drawing
{
	
	/// <summary> The data for this blip store entry.  Typically this is the raw image data</summary>
	class BlipStoreEntry:EscherAtom
	{
		/// <summary> Accessor for the blip type
		/// 
		/// </summary>
		/// <returns> 
		/// </returns>
		virtual public BlipType BlipType
		{
			get
			{
				return type;
			}
			
		}
		/// <summary> Accessor for the reference count on the blip
		/// 
		/// </summary>
		/// <returns> the reference count on the blip
		/// </returns>
		virtual internal int ReferenceCount
		{
			get
			{
				return referenceCount;
			}
			
		}
		/// <summary> Accessor for the image data.  
		/// 
		/// </summary>
		/// <returns> the image data
		/// </returns>
		virtual internal sbyte[] ImageData
		{
			get
			{
				sbyte[] allData = Bytes;
				sbyte[] imageData = new sbyte[allData.Length - IMAGE_DATA_OFFSET];
				Array.Copy(allData, IMAGE_DATA_OFFSET, imageData, 0, imageData.Length);
				return imageData;
			}
			
		}
		/// <summary> The type of the blip</summary>
		private BlipType type;
		
		
		/// <summary> The image data read in</summary>
		new private sbyte[] data;
		
		/// <summary> The .Length of the image data</summary>
		private int imageDataLength;
		
		/// <summary> The reference count on this blip</summary>
		private int referenceCount;
		
		/// <summary> Flag to indicate that this entry was specified by the API, and not
		/// read in
		/// </summary>
		private bool write;
		
		/// <summary> The start of the image data within this blip entry</summary>
		private const int IMAGE_DATA_OFFSET = 61;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="">erd
		/// </param>
		public BlipStoreEntry(EscherRecordData erd):base(erd)
		{
			type = BlipType.getType(Instance);
			write = false;
			sbyte[] bytes = Bytes;
			referenceCount = IntegerHelper.getInt(bytes[24], bytes[25], bytes[26], bytes[27]);
		}
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="d">the drawing
		/// </param>
		/// <exception cref=""> IOException
		/// </exception>
		public BlipStoreEntry(Drawing d):base(EscherRecordType.BSE)
		{
			type = BlipType.PNG;
			Version = 2;
			Instance = type.Value;
			
			sbyte[] imageData = d.ImageBytes;
			imageDataLength = imageData.Length;
			data = new sbyte[imageDataLength + IMAGE_DATA_OFFSET];
			Array.Copy(imageData, 0, data, IMAGE_DATA_OFFSET, imageDataLength);
			referenceCount = d.ReferenceCount;
			write = true;
		}
		
		/// <summary> Gets the data for this blip so that it can be written out
		/// 
		/// </summary>
		/// <returns> the data for the blip
		/// </returns>
		/// <summary> Gets the data for this blip so that it can be written out
		/// </summary>
		/// <returns> the data for the blip
		/// </returns>
		public override sbyte[] Data
		{
		get
		{
		if (write)
		{
		// Drawing has been specified by API
		
		// Type on win32
		data[0] = (sbyte) type.Value;
		
		// Type on MacOs
		data[1] = (sbyte) type.Value;
		
		// The blip identifier
		//    IntegerHelper.getTwoBytes(0xfce1, data, 2);
		
		// Unused tags - 18 bytes
		//    System.Array.Copy(stuff, 0, data, 2, stuff.Length);
		
		// The size of the file
		IntegerHelper.getFourBytes(imageDataLength + 8 + 17, data, 20);
		
		// The reference count on the blip
		IntegerHelper.getFourBytes(referenceCount, data, 24);
		
		// Offset in the delay stream
		IntegerHelper.getFourBytes(0, data, 28);
		
		// Usage byte
		data[32] = (sbyte) 0;
		
		// Length of the blip name
		data[33] = (sbyte) 0;
		
		// Last two bytes unused
		data[34] = (sbyte) 0x7e;
		data[35] = (sbyte) 0x01;
		
		// The blip itself
		data[36] = (sbyte) 0;
		data[37] = (sbyte) 0x6e;
		
		// The blip identifier
		IntegerHelper.getTwoBytes(0xf01e, data, 38);
		
		// The .Length of the blip.  This is the .Length of the image file plus 
		// 16 bytes
		IntegerHelper.getFourBytes(imageDataLength + 17, data, 40);
		
		// Unknown stuff
		//    System.Array.Copy(stuff, 0, data, 44, stuff.Length);
		}
		else
		{
		// drawing has been read in
		data = Bytes;
		}
		
		return setHeaderData(data);
		}
		}
		
		
		
		/// <summary> Reduces the reference count in this blip.  Called when a drawing is
		/// removed
		/// </summary>
		internal virtual void  dereference()
		{
			referenceCount--;
			Assert.verify(referenceCount >= 0);
		}
	}
}
