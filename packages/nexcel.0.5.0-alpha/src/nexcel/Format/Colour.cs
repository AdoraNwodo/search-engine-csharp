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
namespace NExcel.Format
{
	
	/// <summary> Enumeration class which contains the various colours available within
	/// the standard Excel colour palette
	/// 
	/// </summary>
	public class Colour
	{
		/// <summary> Gets the value of this colour.  This is the value that is written to 
		/// the generated Excel file
		/// 
		/// </summary>
		/// <returns> the binary value
		/// </returns>
		virtual public int Value
		{
			get
			{
				return value_Renamed;
			}
			
		}
		/// <summary> Gets the string description for display purposes
		/// 
		/// </summary>
		/// <returns> the string description
		/// </returns>
		virtual public string Description
		{
			get
			{
				return string_Renamed;
			}
			
		}
		/// <summary> Gets the default red content of this colour.  Used when writing the
		/// default colour palette
		/// 
		/// </summary>
		/// <returns> the red content of this colour
		/// </returns>
		virtual public int DefaultRed
		{
			get
			{
				return red;
			}
			
		}
		/// <summary> Gets the default green content of this colour.  Used when writing the
		/// default colour palette
		/// 
		/// </summary>
		/// <returns> the green content of this colour
		/// </returns>
		virtual public int DefaultGreen
		{
			get
			{
				return green;
			}
			
		}
		/// <summary> Gets the default blue content of this colour.  Used when writing the
		/// default colour palette
		/// 
		/// </summary>
		/// <returns> the blue content of this colour
		/// </returns>
		virtual public int DefaultBlue
		{
			get
			{
				return blue;
			}
			
		}
		/// <summary> Returns all available colours - used when generating the default palette
		/// 
		/// </summary>
		/// <returns> all available colours
		/// </returns>
		public static Colour[] AllColours
		{
			get
			{
				return colours;
			}
			
		}
		/// <summary> The internal numerical representation of the colour</summary>
		private int value_Renamed;
		
		/// <summary> The default "red" value</summary>
		private int red;
		
		/// <summary> The default "green" value</summary>
		private int green;
		
		/// <summary> The default "blue" value</summary>
		private int blue;
		
		/// <summary> The display string for the colour.  Used when presenting the 
		/// format information
		/// </summary>
		private string string_Renamed;
		
		private bool initialized;
		
		/// <summary> The list of internal colours</summary>
		private static Colour[] colours;
		
		/// <summary> Private constructor
		/// 
		/// </summary>
		/// <param name="val">
		/// </param>
		/// <param name="s">the display string
		/// </param>
		/// <param name="r">the default red value
		/// </param>
		/// <param name="g">the default green value
		/// </param>
		/// <param name="b">the default blue value
		/// </param>
		protected internal Colour(int val, string s, int r, int g, int b)
		{
			value_Renamed = val;
			string_Renamed = s;
			red = r;
			green = g;
			blue = b;
			
			Colour[] oldcols = colours;
			colours = new Colour[oldcols.Length + 1];
			Array.Copy((System.Array) oldcols, 0, (System.Array) colours, 0, oldcols.Length);
			colours[oldcols.Length] = this;
			initialized = true;
		}
		
		/// <summary> Gets the internal colour from the value
		/// 
		/// </summary>
		/// <param name="val">
		/// </param>
		/// <returns> the colour with that value
		/// </returns>
		public static Colour getInternalColour(int val)
		{
			for (int i = 0; i < colours.Length; i++)
			{
				if (colours[i].Value == val)
				{
					return colours[i];
				}
			}
			
			return UNKNOWN;
		}
		
		// Major colours
		public static readonly Colour UNKNOWN = new Colour(0x7fee, "unknown", 0, 0, 0);
		public static readonly Colour BLACK = new Colour(0x7fff, "black", 0, 0, 0);
		public static readonly Colour WHITE = new Colour(0x9, "white", 0xff, 0xff, 0xff);
		public static readonly Colour DEFAULT_BACKGROUND1 = new Colour(0x0, "default background", 0xff, 0xff, 0xff);
		public static readonly Colour DEFAULT_BACKGROUND = new Colour(0xc0, "default background", 0xff, 0xff, 0xff);
		public static readonly Colour PALETTE_BLACK = new Colour(0x8, "black", 0x1, 0, 0);
		// the first item in the colour palette
		
		// Other colours
		public static readonly Colour RED = new Colour(0xa, "red", 0xff, 0, 0);
		public static readonly Colour BRIGHT_GREEN = new Colour(0xb, "bright green", 0, 0xff, 0);
		public static readonly Colour BLUE = new Colour(0xc, "blue", 0, 0, 0xff);
		public static readonly Colour YELLOW = new Colour(0xd, "yellow", 0xff, 0xff, 0);
		public static readonly Colour PINK = new Colour(0xe, "pink", 0xff, 0, 0xff);
		public static readonly Colour TURQUOISE = new Colour(0xf, "turquoise", 0, 0xff, 0xff);
		public static readonly Colour DARK_RED = new Colour(0x10, "dark red", 0x80, 0, 0);
		public static readonly Colour GREEN = new Colour(0x11, "green", 0, 0x80, 0);
		public static readonly Colour DARK_BLUE = new Colour(0x12, "dark blue", 0, 0, 0x80);
		public static readonly Colour DARK_YELLOW = new Colour(0x13, "dark yellow", 0x80, 0x80, 0);
		public static readonly Colour VIOLET = new Colour(0x14, "violet", 0x80, 0x80, 0);
		public static readonly Colour TEAL = new Colour(0x15, "teal", 0, 0x80, 0x80);
		public static readonly Colour GREY_25_PERCENT = new Colour(0x16, "grey 25%", 0xc0, 0xc0, 0xc0);
		public static readonly Colour GREY_50_PERCENT = new Colour(0x17, "grey 50%", 0x80, 0x80, 0x80);
		public static readonly Colour PERIWINKLE = new Colour(0x18, "periwinkle%", 0x99, 0x99, 0xff);
		public static readonly Colour PLUM2 = new Colour(0x19, "plum", 0x99, 0x33, 0x66);
		public static readonly Colour IVORY = new Colour(0x1a, "ivory", 0xff, 0xff, 0xcc);
		public static readonly Colour LIGHT_TURQUOISE2 = new Colour(0x1b, "light turquoise", 0xcc, 0xff, 0xff);
		public static readonly Colour DARK_PURPLE = new Colour(0x1c, "dark purple", 0x66, 0x0, 0x66);
		public static readonly Colour CORAL = new Colour(0x1d, "coral", 0xff, 0x80, 0x80);
		public static readonly Colour OCEAN_BLUE = new Colour(0x1e, "ocean blue", 0x0, 0x66, 0xcc);
		public static readonly Colour ICE_BLUE = new Colour(0x1f, "ice blue", 0xcc, 0xcc, 0xff);
		public static readonly Colour DARK_BLUE2 = new Colour(0x20, "dark blue", 0, 0, 0x80);
		public static readonly Colour PINK2 = new Colour(0x21, "pink", 0xff, 0, 0xff);
		public static readonly Colour YELLOW2 = new Colour(0x22, "yellow", 0xff, 0xff, 0x0);
		public static readonly Colour TURQOISE2 = new Colour(0x23, "turqoise", 0x0, 0xff, 0xff);
		public static readonly Colour VIOLET2 = new Colour(0x24, "violet", 0x80, 0x0, 0x80);
		public static readonly Colour DARK_RED2 = new Colour(0x25, "dark red", 0x80, 0x0, 0x0);
		public static readonly Colour TEAL2 = new Colour(0x26, "teal", 0x0, 0x80, 0x80);
		public static readonly Colour BLUE2 = new Colour(0x27, "blue", 0x0, 0x0, 0xff);
		public static readonly Colour SKY_BLUE = new Colour(0x28, "sky blue", 0, 0xcc, 0xff);
		public static readonly Colour LIGHT_TURQUOISE = new Colour(0x29, "light turquoise", 0xcc, 0xff, 0xff);
		public static readonly Colour LIGHT_GREEN = new Colour(0x2a, "light green", 0xcc, 0xff, 0xcc);
		public static readonly Colour VERY_LIGHT_YELLOW = new Colour(0x2b, "very light yellow", 0xff, 0xff, 0x99);
		public static readonly Colour PALE_BLUE = new Colour(0x2c, "pale blue", 0x99, 0xcc, 0xff);
		public static readonly Colour ROSE = new Colour(0x2d, "rose", 0xff, 0x99, 0xcc);
		public static readonly Colour LAVENDER = new Colour(0x2e, "lavender", 0xcc, 0x99, 0xff);
		public static readonly Colour TAN = new Colour(0x2f, "tan", 0xff, 0xcc, 0x99);
		public static readonly Colour LIGHT_BLUE = new Colour(0x30, "light blue", 0x33, 0x66, 0xff);
		public static readonly Colour AQUA = new Colour(0x31, "aqua", 0x33, 0xcc, 0xcc);
		public static readonly Colour LIME = new Colour(0x32, "lime", 0x99, 0xcc, 0);
		public static readonly Colour GOLD = new Colour(0x33, "gold", 0xff, 0xcc, 0);
		public static readonly Colour LIGHT_ORANGE = new Colour(0x34, "light orange", 0xff, 0x99, 0);
		public static readonly Colour ORANGE = new Colour(0x35, "orange", 0xff, 0x66, 0);
		public static readonly Colour BLUE_GREY = new Colour(0x36, "blue grey", 0x66, 0x66, 0xcc);
		public static readonly Colour GREY_40_PERCENT = new Colour(0x37, "grey 40%", 0x96, 0x96, 0x96);
		public static readonly Colour DARK_TEAL = new Colour(0x38, "dark teal", 0, 0x33, 0x66);
		public static readonly Colour SEA_GREEN = new Colour(0x39, "sea green", 0x33, 0x99, 0x66);
		public static readonly Colour DARK_GREEN = new Colour(0x3a, "dark green", 0, 0x33, 0);
		public static readonly Colour OLIVE_GREEN = new Colour(0x3b, "olive green", 0x33, 0x33, 0);
		public static readonly Colour BROWN = new Colour(0x3c, "brown", 0x99, 0x33, 0);
		public static readonly Colour PLUM = new Colour(0x3d, "plum", 0x99, 0x33, 0x66);
		public static readonly Colour INDIGO = new Colour(0x3e, "indigo", 0x33, 0x33, 0x99);
		public static readonly Colour GREY_80_PERCENT = new Colour(0x3f, "grey 80%", 0x33, 0x33, 0x33);
		
		// Colours added for backwards compatibility
		public static readonly Colour GRAY_80 = GREY_80_PERCENT;
		public static readonly Colour GRAY_50 = GREY_50_PERCENT;
		public static readonly Colour GRAY_25 = GREY_25_PERCENT;
		static Colour()
		{
			colours = new Colour[0];
		}
	}
}