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
using NExcel.Read.Biff;
using NExcel.Format;
namespace NExcel.Biff
{
	
	/// <summary> A record representing the RGB colour palette</summary>
	public class PaletteRecord:WritableRecordData
	{
		private void  InitBlock()
		{
			rgbColours = new RGB[numColours];
		}
		/// <summary> Accessor for the dirty flag, which indicates if this palette has been
		/// modified
		/// 
		/// </summary>
		/// <returns> TRUE if the palette has been modified, FALSE if it is the default
		/// </returns>
		virtual public bool Dirty
		{
			get
			{
				return dirty;
			}
			
		}
		/// <summary> The internal RGB structure</summary>
		private class RGB
		{
			/// <summary> The red component of this colour</summary>
			internal int red;
			
			/// <summary> The green component of this colour</summary>
			internal int green;
			
			/// <summary> The blue component of this colour</summary>
			internal int blue;
			
			/// <summary> Constructor
			/// 
			/// </summary>
			/// <param name="r">the red component
			/// </param>
			/// <param name="g">the green component
			/// </param>
			/// <param name="b">the blue component
			/// </param>
			internal RGB(int r, int g, int b)
			{
				red = r;
				green = g;
				blue = b;
			}
		}
		
		/// <summary> The list of bespoke rgb colours used by this sheet</summary>
		private RGB[] rgbColours;
		
		/// <summary> A dirty flag indicating that this palette has been tampered with
		/// in some way
		/// </summary>
		private bool dirty;
		
		/// <summary> Flag indicating that the palette was read in</summary>
		private bool read;
		
		/// <summary> Initialized flag</summary>
		private bool initialized;
		
		/// <summary> The number of colours in the palette</summary>
		private const int numColours = 56;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="t">the raw bytes
		/// </param>
		public PaletteRecord(Record t):base(t)
		{
			InitBlock();
			
			initialized = false;
			dirty = false;
			read = true;
		}
		
		/// <summary> Default constructor - used when there is no palette specified</summary>
		public PaletteRecord():base(NExcel.Biff.Type.PALETTE)
		{
			InitBlock();
			
			initialized = true;
			dirty = false;
			read = false;
			
			// Initialize the array with all the default colours
			Colour[] colours = Colour.AllColours;
			
			for (int i = 0; i < colours.Length; i++)
			{
				Colour c = colours[i];
				setColourRGB(c, c.DefaultRed, c.DefaultGreen, c.DefaultBlue);
			}
		}
		
		/// <summary> Accessor for the binary data - used when copying
		/// 
		/// </summary>
		/// <returns> the binary data
		/// </returns>
		public override sbyte[] getData()
		{
			// Palette was read in, but has not been changed
			if (read && !dirty)
			{
				return getRecord().Data;
			}
			
			sbyte[] data = new sbyte[numColours * 4 + 2];
			int pos = 0;
			
			// Set the number of records
			IntegerHelper.getTwoBytes(numColours, data, pos);
			
			// Set the rgb content
			for (int i = 0; i < numColours; i++)
			{
				pos = i * 4 + 2;
				data[pos] = (sbyte) rgbColours[i].red;
				data[pos + 1] = (sbyte) rgbColours[i].green;
				data[pos + 2] = (sbyte) rgbColours[i].blue;
			}
			
			return data;
		}
		
		/// <summary> Initialize the record data</summary>
		private void  initialize()
		{
			sbyte[] data = getRecord().Data;
			
			int numrecords = IntegerHelper.getInt(data[0], data[1]);
			
			for (int i = 0; i < numrecords; i++)
			{
				int pos = i * 4 + 2;
				int red = IntegerHelper.getInt(data[pos], (sbyte) 0);
				int green = IntegerHelper.getInt(data[pos + 1], (sbyte) 0);
				int blue = IntegerHelper.getInt(data[pos + 2], (sbyte) 0);
				rgbColours[i] = new RGB(red, green, blue);
			}
			
			initialized = true;
		}
		
		/// <summary> Sets the RGB value for the specified colour for this workbook
		/// 
		/// </summary>
		/// <param name="c">the colour whose RGB value is to be overwritten
		/// </param>
		/// <param name="r">the red portion to set (0-255)
		/// </param>
		/// <param name="g">the green portion to set (0-255)
		/// </param>
		/// <param name="b">the blue portion to set (0-255)
		/// </param>
		public virtual void  setColourRGB(Colour c, int r, int g, int b)
		{
			// Only colours on the standard palette with values 8-64 are acceptable
			int pos = c.Value - 8;
			if (pos < 0 || pos >= numColours)
			{
				return ;
			}
			
			if (!initialized)
			{
				initialize();
			}
			
			// Force the colours into the range 0-255
			r = setValueRange(r, 0, 0xff);
			g = setValueRange(g, 0, 0xff);
			b = setValueRange(b, 0, 0xff);
			
			rgbColours[pos] = new RGB(r, g, b);
			
			// Indicate that the palette has been modified
			dirty = true;
		}
		
		/// <summary> Forces the value passed in to be between the range passed in
		/// 
		/// </summary>
		/// <param name="val">the value to constrain
		/// </param>
		/// <param name="min">the minimum acceptable value
		/// </param>
		/// <param name="max">the maximum acceptable value
		/// </param>
		/// <returns> the constrained value
		/// </returns>
		private int setValueRange(int val, int min, int max)
		{
			val = System.Math.Min(val, min);
			val = System.Math.Min(val, max);
			return val;
		}
	}
}
