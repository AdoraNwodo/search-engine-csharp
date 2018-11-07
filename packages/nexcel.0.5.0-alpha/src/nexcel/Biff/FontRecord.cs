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
using NExcel;
using NExcel.Read.Biff;
using NExcel.Format;
namespace NExcel.Biff
{
	
	/// <summary> A record containing the necessary data for the font information</summary>
	public class FontRecord:WritableRecordData, Font
	{
		/// <summary> Accessor to see whether this object is initialized or not.
		/// 
		/// </summary>
		/// <returns> TRUE if this font record has been initialized, FALSE otherwise
		/// </returns>
//		virtual public bool Initialized
//		{
//			get
//			{
//				return initialized;
//			}
//			
//		}
		virtual public bool IsInitialized()
		{
				return initialized;
		}

		
		/// <summary> Accessor for the font index
		/// 
		/// </summary>
		/// <returns> the font index
		/// </returns>
		virtual public int FontIndex
		{
			get
			{
				return fontIndex;
			}
			
		}
		/// <summary> Sets the point size for this font, if the font hasn't been initialized
		/// 
		/// </summary>
		/// <param name="ps">the point size
		/// </param>
		virtual protected internal int FontPointSize
		{
			set
			{
				Assert.verify(!initialized);
				
				pointHeight = value;
			}
			
		}
		/// <summary> Gets the point size for this font, if the font hasn't been initialized
		/// 
		/// </summary>
		/// <returns> the point size
		/// </returns>
		virtual public int PointSize
		{
			get
			{
				return pointHeight;
			}
			
		}
		/// <summary> Sets the bold style for this font, if the font hasn't been initialized
		/// 
		/// </summary>
		/// <param name="bs">the bold style
		/// </param>
		virtual protected internal int FontBoldStyle
		{
			set
			{
				Assert.verify(!initialized);
				
				boldWeight = value;
			}
			
		}
		/// <summary> Gets the bold weight for this font
		/// 
		/// </summary>
		/// <returns> the bold weight for this font
		/// </returns>
		virtual public int BoldWeight
		{
			get
			{
				return boldWeight;
			}
			
		}
		/// <summary> Sets the italic indicator for this font, if the font hasn't been
		/// initialized
		/// 
		/// </summary>
		/// <param name="i">the italic flag
		/// </param>
		virtual protected internal bool FontItalic
		{
			set
			{
				Assert.verify(!initialized);
				
				italic = value;
			}
			
		}
		/// <summary> Returns the italic flag
		/// 
		/// </summary>
		/// <returns> TRUE if this font is italic, FALSE otherwise
		/// </returns>
		virtual public bool Italic
		{
			get
			{
				return italic;
			}
			
		}
		/// <summary> Sets the underline style for this font, if the font hasn't been
		/// initialized
		/// 
		/// </summary>
		/// <param name="us">the underline style
		/// </param>
		virtual protected internal int FontUnderlineStyle
		{
			set
			{
				Assert.verify(!initialized);
				
				underlineStyle = value;
			}
			
		}
		/// <summary> Gets the underline style for this font
		/// 
		/// </summary>
		/// <returns> the underline style
		/// </returns>
		virtual public UnderlineStyle UnderlineStyle
		{
			get
			{
				return UnderlineStyle.getStyle(underlineStyle);
			}
			
		}
		/// <summary> Sets the colour for this font, if the font hasn't been
		/// initialized
		/// 
		/// </summary>
		/// <param name="c">the colour
		/// </param>
		virtual protected internal int FontColour
		{
			set
			{
				Assert.verify(!initialized);
				
				colourIndex = value;
			}
			
		}
		/// <summary> Gets the colour for this font
		/// 
		/// </summary>
		/// <returns> the colour
		/// </returns>
		virtual public Colour Colour
		{
			get
			{
				return Colour.getInternalColour(colourIndex);
			}
			
		}
		/// <summary> Sets the script style (eg. superscript, subscript) for this font,
		/// if the font hasn't been initialized
		/// 
		/// </summary>
		/// <param name="ss">the colour
		/// </param>
		virtual protected internal int FontScriptStyle
		{
			set
			{
				Assert.verify(!initialized);
				
				scriptStyle = value;
			}
			
		}
		/// <summary> Gets the script style
		/// 
		/// </summary>
		/// <returns> the script style
		/// </returns>
		virtual public ScriptStyle ScriptStyle
		{
			get
			{
				return ScriptStyle.getStyle(scriptStyle);
			}
			
		}
		/// <summary> Gets the name of this font
		/// 
		/// </summary>
		/// <returns> the name of this font
		/// </returns>
		virtual public string Name
		{
			get
			{
				return name;
			}
			
		}
		/// <summary> Accessor for the strike out flag
		/// 
		/// </summary>
		/// <returns> TRUE if this font is struck out, FALSE otherwise
		/// </returns>
		virtual public bool Struckout
		{
			get
			{
				return struckout;
			}
			
		}
		/// <summary> Sets the struck out flag
		/// 
		/// </summary>
		/// <param name="os">TRUE if the font is struck out, false otherwise
		/// </param>
		virtual protected internal bool FontStruckout
		{
			set
			{
				struckout = value;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The point height of this font</summary>
		private int pointHeight;
		/// <summary> The index into the colour palette</summary>
		private int colourIndex;
		/// <summary> The bold weight for this font (normal or bold)</summary>
		private int boldWeight;
		/// <summary> The style of the script (italic or normal)</summary>
		private int scriptStyle;
		/// <summary> The underline style for this font (none, single, double etc)</summary>
		private int underlineStyle;
		/// <summary> The font family</summary>
		private sbyte fontFamily;
		/// <summary> The character set</summary>
		private sbyte characterSet;
		
		/// <summary> Indicates whether or not this font is italic</summary>
		private bool italic;
		/// <summary> Indicates whether or not this font is struck out</summary>
		private bool struckout;
		/// <summary> The name of this font</summary>
		private string name;
		/// <summary> Flag to indicate whether the derived data (such as the font index) has
		/// been initialized or not
		/// </summary>
		private bool initialized;
		
		/// <summary> The index of this font in the font list</summary>
		private int fontIndex;
		
		/// <summary> Dummy indicators for overloading the constructor</summary>
		public class Biff7
		{
		}
		
		public static readonly Biff7 biff7 = new Biff7();
		
		/// <summary> The conversion factor between microsoft internal units and point size</summary>
		private const int EXCEL_UNITS_PER_POINT = 20;
		
		/// <summary> Constructor, used when creating a new font for writing out.
		/// 
		/// </summary>
		/// <param name="bold">the bold indicator
		/// </param>
		/// <param name="ps">the point size
		/// </param>
		/// <param name="us">the underline style
		/// </param>
		/// <param name="fn">the name
		/// </param>
		/// <param name="it">italicised indicator
		/// </param>
		/// <param name="ss">the script style
		/// </param>
		/// <param name="ci">the colour index
		/// </param>
		protected internal FontRecord(string fn, int ps, int bold, bool it, int us, int ci, int ss):base(NExcel.Biff.Type.FONT)
		{
			boldWeight = bold;
			underlineStyle = us;
			name = fn;
			pointHeight = ps;
			italic = it;
			scriptStyle = ss;
			colourIndex = ci;
			initialized = false;
			struckout = false;
		}
		
		/// <summary> Constructs this object from the raw data.  Used when reading in a
		/// format record
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public FontRecord(Record t, WorkbookSettings ws):base(t)
		{
			
			sbyte[] data = getRecord().Data;
			
			pointHeight = IntegerHelper.getInt(data[0], data[1]) / EXCEL_UNITS_PER_POINT;
			colourIndex = IntegerHelper.getInt(data[4], data[5]);
			boldWeight = IntegerHelper.getInt(data[6], data[7]);
			scriptStyle = IntegerHelper.getInt(data[8], data[9]);
			underlineStyle = data[10];
			fontFamily = data[11];
			characterSet = data[12];
			initialized = false;
			
			if ((data[2] & 0x02) != 0)
			{
				italic = true;
			}
			
			if ((data[2] & 0x08) != 0)
			{
				struckout = true;
			}
			
			int numChars = data[14];
			if (data[15] == 0)
			{
				name = StringHelper.getString(data, numChars, 16, ws);
			}
			else if (data[15] == 1)
			{
				name = StringHelper.getUnicodeString(data, numChars, 16);
			}
			else
			{
				// Some font names don't have the unicode indicator
				name = StringHelper.getString(data, numChars, 15, ws);
			}
		}
		
		/// <summary> Constructs this object from the raw data.  Used when reading in a
		/// format record
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <param name="dummy">dummy overload
		/// </param>
		public FontRecord(Record t, WorkbookSettings ws, Biff7 dummy):base(t)
		{
			
			sbyte[] data = getRecord().Data;
			
			pointHeight = IntegerHelper.getInt(data[0], data[1]) / EXCEL_UNITS_PER_POINT;
			colourIndex = IntegerHelper.getInt(data[4], data[5]);
			boldWeight = IntegerHelper.getInt(data[6], data[7]);
			scriptStyle = IntegerHelper.getInt(data[8], data[9]);
			underlineStyle = data[10];
			fontFamily = data[11];
			initialized = false;
			
			if ((data[2] & 0x02) != 0)
			{
				italic = true;
			}
			
			if ((data[2] & 0x08) != 0)
			{
				struckout = true;
			}
			
			int numChars = data[14];
			name = StringHelper.getString(data, numChars, 15, ws);
		}
		
		/// <summary> Publicly available copy constructor
		/// 
		/// </summary>
		/// <param name="f">the font to copy
		/// </param>
		protected internal FontRecord(Font f):base(NExcel.Biff.Type.FONT)
		{
			
			Assert.verify(f != null);
			
			pointHeight = f.PointSize;
			colourIndex = f.Colour.Value;
			boldWeight = f.BoldWeight;
			scriptStyle = f.ScriptStyle.Value;
			underlineStyle = f.UnderlineStyle.Value;
			italic = f.Italic;
			name = f.Name;
			struckout = false;
			initialized = false;
		}
		
		/// <summary> Gets the byte data for writing out
		/// 
		/// </summary>
		/// <returns> the raw data
		/// </returns>
		public override sbyte[] getData()
		{
			sbyte[] data = new sbyte[16 + name.Length * 2];
			
			// Excel expects font heights in 1/20ths of a point
			IntegerHelper.getTwoBytes(pointHeight * EXCEL_UNITS_PER_POINT, data, 0);
			
			// Set the font attributes to be zero for now
			if (italic)
			{
				data[2] |= 0x2;
			}
			
			if (struckout)
			{
				data[2] |= 0x08;
			}
			
			// Set the index to the colour palette
			IntegerHelper.getTwoBytes(colourIndex, data, 4);
			
			// Bold style
			IntegerHelper.getTwoBytes(boldWeight, data, 6);
			
			// Script style
			IntegerHelper.getTwoBytes(scriptStyle, data, 8);
			
			// Underline style
			data[10] = (sbyte) underlineStyle;
			
			// Set the font family to be 0
			data[11] = fontFamily;
			
			// Set the character set to be zero
			data[12] = characterSet;
			
			// Set the reserved bit to be zero
			data[13] = 0;
			
			// Set the .Length of the font name
			data[14] = (sbyte) name.Length;
			
			data[15] = (sbyte) 1;
			
			// Copy in the string
			StringHelper.getUnicodeBytes(name, data, 16);
			
			return data;
		}
		
		/// <summary> Sets the font index of this record.  Called from the FormattingRecords
		/// object
		/// 
		/// </summary>
		/// <param name="pos">the position of this font in the workbooks font list
		/// </param>
		public void  initialize(int pos)
		{
			fontIndex = pos;
			initialized = true;
		}
		
		/// <summary> Resets the initialize flag.  This is called by the constructor of
		/// WritableWorkbookImpl to reset the statically declared fonts
		/// </summary>
		public void  uninitialize()
		{
			initialized = false;
		}
		
		/// <summary> Standard hash code method</summary>
		/// <returns> the hash code for this object
		/// </returns>
		public override int GetHashCode()
		{
			return name.GetHashCode();
		}
		
		/// <summary> Standard equals method</summary>
		/// <param name="o">the object to compare
		/// </param>
		/// <returns> TRUE if the objects are equal, FALSE otherwise
		/// </returns>
		public  override bool Equals(System.Object o)
		{
			if (o == this)
			{
				return true;
			}
			
			if (!(o is FontRecord))
			{
				return false;
			}
			
			FontRecord font = (FontRecord) o;
			
			if (pointHeight == font.pointHeight && colourIndex == font.colourIndex && boldWeight == font.boldWeight && scriptStyle == font.scriptStyle && underlineStyle == font.underlineStyle && italic == font.italic && struckout == font.struckout && fontFamily == font.fontFamily && characterSet == font.characterSet && name.Equals(font.name))
			{
				return true;
			}
			
			return false;
		}
		static FontRecord()
		{
			logger = Logger.getLogger(typeof(FontRecord));
		}
	}
}
