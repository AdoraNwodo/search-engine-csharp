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
using NExcelUtils;
using common;
using NExcel.Format;
using NExcel.Read.Biff;
namespace NExcel.Biff
{
	
	/// <summary> Holds an extended formatting record</summary>
	public class XFRecord:WritableRecordData, NExcel.Format.CellFormat
	{
		/// <summary> Gets the java date format for this format record
		/// 
		/// </summary>
		/// <returns> returns the date format
		/// </returns>
		virtual public DateTimeFormatInfo DateFormat
		{
			get
			{
				return dateFormat;
			}
			
		}
		/// <summary> Gets the java number format for this format record
		/// 
		/// </summary>
		/// <returns> returns the number format
		/// </returns>
		virtual public NumberFormatInfo NumberFormat
		{
			get
			{
				return numberFormat;
			}
			
		}
		/// <summary> Gets the lookup number of the format record
		/// 
		/// </summary>
		/// <returns> returns the lookup number of the format record
		/// </returns>
		virtual public int FormatRecord
		{
			get
			{
				return formatIndex;
			}
			
		}
		/// <summary> Sees if this format is a date format
		/// 
		/// </summary>
		/// <returns> TRUE if this refers to a built in date format
		/// </returns>
		virtual public bool isDate()
		{
				return date;
		}
		/// <summary> Sees if this format is a number format
		/// 
		/// </summary>
		/// <returns> TRUE if this refers to a built in date format
		/// </returns>
		virtual public bool isNumber()
		{
				return number;
		}
		/// <summary> Accessor for the hidden flag
		/// 
		/// </summary>
		/// <returns> TRUE if this XF record hides the cell, FALSE otherwise
		/// </returns>
		virtual protected internal bool Hidden
		{
			get
			{
				return hidden;
			}
			
		}
		/// <summary> Sets whether or not this XF record locks the cell
		/// 
		/// </summary>
		/// <param name="l">the locked flag
		/// </param>
		virtual protected internal bool XFLocked
		{
			set
			{
				locked = value;
			}
			
		}
		/// <summary> Sets the cell options
		/// 
		/// </summary>
		/// <param name="opt">the cell options
		/// </param>
		virtual protected internal int XFCellOptions
		{
			set
			{
				options |= value;
			}
			
		}
		/// <summary> Sets the horizontal alignment for the data in this cell.
		/// This method should only be called from its writable subclass
		/// CellXFRecord
		/// 
		/// </summary>
		/// <param name="a">the alignment
		/// </param>
		virtual protected internal Alignment XFAlignment
		{
			set
			{
				Assert.verify(!initialized);
				align = value;
			}
			
		}
		/// <summary> Sets the shrink to fit flag
		/// 
		/// </summary>
		/// <param name="s">the shrink to fit flag
		/// </param>
		virtual protected internal bool XFShrinkToFit
		{
			set
			{
				Assert.verify(!initialized);
				shrinkToFit = value;
			}
			
		}
		/// <summary> Gets the horizontal cell alignment
		/// 
		/// </summary>
		/// <returns> the alignment
		/// </returns>
		virtual public Alignment Alignment
		{
			get
			{
				if (!formatInfoInitialized)
				{
					initializeFormatInformation();
				}
				
				return align;
			}
			
		}
		/// <summary> Gets the shrink to fit flag
		/// 
		/// </summary>
		/// <returns> TRUE if this format is shrink to fit, FALSE otherise
		/// </returns>
		virtual public bool ShrinkToFit
		{
			get
			{
				if (!formatInfoInitialized)
				{
					initializeFormatInformation();
				}
				
				return shrinkToFit;
			}
			
		}
		/// <summary> Gets the vertical cell alignment
		/// 
		/// </summary>
		/// <returns> the alignment
		/// </returns>
		virtual public VerticalAlignment VerticalAlignment
		{
			get
			{
				if (!formatInfoInitialized)
				{
					initializeFormatInformation();
				}
				
				return valign;
			}
			
		}
		/// <summary> Gets the orientation
		/// 
		/// </summary>
		/// <returns> the orientation
		/// </returns>
		virtual public Orientation Orientation
		{
			get
			{
				if (!formatInfoInitialized)
				{
					initializeFormatInformation();
				}
				
				return orientation;
			}
			
		}
		/// <summary> Gets the background colour used by this cell
		/// 
		/// </summary>
		/// <returns> the foreground colour
		/// </returns>
		virtual public Colour BackgroundColour
		{
			get
			{
				if (!formatInfoInitialized)
				{
					initializeFormatInformation();
				}
				
				return backgroundColour;
			}
			
		}
		/// <summary> Gets the pattern used by this cell format
		/// 
		/// </summary>
		/// <returns> the background pattern
		/// </returns>
		virtual public Pattern Pattern
		{
			get
			{
				if (!formatInfoInitialized)
				{
					initializeFormatInformation();
				}
				
				return pattern;
			}
			
		}
		/// <summary> Sets the vertical alignment for the data in this cell
		/// This method should only be called from its writable subclass
		/// CellXFRecord
		/// 
		/// </summary>
		/// <param name="va">the vertical alignment
		/// </param>
		virtual protected internal VerticalAlignment XFVerticalAlignment
		{
			set
			{
				Assert.verify(!initialized);
				valign = value;
			}
			
		}
		/// <summary> Sets the vertical alignment for the data in this cell
		/// This method should only be called from its writable subclass
		/// CellXFRecord
		/// 
		/// </summary>
		/// <param name="o">the orientation
		/// </param>
		virtual protected internal Orientation XFOrientation
		{
			set
			{
				Assert.verify(!initialized);
				orientation = value;
			}
			
		}
		/// <summary> Sets whether the data in this cell is wrapped
		/// This method should only be called from its writable subclass
		/// CellXFRecord
		/// 
		/// </summary>
		/// <param name="w">the wrap flag
		/// </param>
		virtual protected internal bool XFWrap
		{
			set
			{
				Assert.verify(!initialized);
				wrap = value;
			}
			
		}
		/// <summary> Gets whether or not the contents of this cell are wrapped
		/// 
		/// </summary>
		/// <returns> TRUE if this cell's contents are wrapped, FALSE otherwise
		/// </returns>
		virtual public bool Wrap
		{
			get
			{
				if (!formatInfoInitialized)
				{
					initializeFormatInformation();
				}
				
				return wrap;
			}
			
		}
		/// <summary> Accessor to see if this format is initialized
		/// 
		/// </summary>
		/// <returns> TRUE if this format is initialized, FALSE otherwise
		/// </returns>
		virtual public bool isInitialized()
		{
				return initialized;
		}
		/// <summary> Accessor to see if this format was read in.  Used when checking merged
		/// cells
		/// 
		/// </summary>
		/// <returns> TRUE if this XF record was read in, FALSE if it was generated by
		/// the user API
		/// </returns>
		virtual public bool Read
		{
			get
			{
				return read;
			}
			
		}
		/// <summary> Gets the format used by this format
		/// 
		/// </summary>
		/// <returns> the format
		/// </returns>
		virtual public NExcel.Format.Format Format
		{
			get
			{
				if (!formatInfoInitialized)
				{
					initializeFormatInformation();
				}
				return excelFormat;
			}
			
		}
		/// <summary> Sets the format index.  This is called during the rationalization process
		/// when some of the duplicate number formats have been removed
		/// </summary>
		/// <param name="newindex">the new format index
		/// </param>
		virtual internal int FormatIndex
		{
			set
			{
				formatIndex = value;
			}
			
		}
		/// <summary> Accessor for the font index.  Called by the FormattingRecords objects
		/// during the rationalization process
		/// </summary>
		/// <returns> the font index
		/// </returns>
		/// <summary> Sets the font index.  This is called during the rationalization process
		/// when some of the duplicate fonts have been removed
		/// </summary>
		/// <param name="newindex">the new index
		/// </param>
		virtual internal int FontIndex
		{
			get
			{
				return fontIndex;
			}
			
			set
			{
				fontIndex = value;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The index to the format record</summary>
		private int formatIndex;
		
		/// <summary> The index of the parent format</summary>
		private int parentFormat;
		
		/// <summary> The format type</summary>
		private XFType xfFormatType;
		
		/// <summary> Indicates whether this is a date formatting record</summary>
		private bool date;
		
		/// <summary> Indicates whether this is a number formatting record</summary>
		private bool number;
		
		/// <summary> The date format for this record.  Deduced when the record is
		/// read in from a spreadsheet
		/// </summary>
		private DateTimeFormatInfo dateFormat;
		
		/// <summary> The number format for this record.  Deduced when the record is read in
		/// from a spreadsheet
		/// </summary>
		private NumberFormatInfo numberFormat;
		
		/// <summary> The used attribute.  Needs to be preserved in order to get accurate
		/// rationalization
		/// </summary>
		private sbyte usedAttributes;
		/// <summary> The index to the font record used by this XF record</summary>
		private int fontIndex;
		/// <summary> Flag to indicate whether this XF record represents a locked cell</summary>
		private bool locked;
		/// <summary> Flag to indicate whether this XF record is hidden</summary>
		private bool hidden;
		/// <summary> The alignment for this cell (left, right, centre)</summary>
		private Alignment align;
		/// <summary> The vertical alignment for the cell (top, bottom, centre)</summary>
		private VerticalAlignment valign;
		/// <summary> The orientation of the cell</summary>
		private Orientation orientation;
		/// <summary> Flag to indicates whether the data (normally text) in the cell will be
		/// wrapped around to fit in the cell width
		/// </summary>
		private bool wrap;
		
		/// <summary> Flag to indicate that this format is shrink to fit</summary>
		private bool shrinkToFit;
		
		/// <summary> The border indicator for the left of this cell</summary>
		private BorderLineStyle leftBorder;
		/// <summary> The border indicator for the right of the cell</summary>
		private BorderLineStyle rightBorder;
		/// <summary> The border indicator for the top of the cell</summary>
		private BorderLineStyle topBorder;
		/// <summary> The border indicator for the bottom of the cell</summary>
		private BorderLineStyle bottomBorder;
		
		/// <summary> The border colour for the left of the cell</summary>
		private Colour leftBorderColour;
		/// <summary> The border colour for the right of the cell</summary>
		private Colour rightBorderColour;
		/// <summary> The border colour for the top of the cell</summary>
		private Colour topBorderColour;
		/// <summary> The border colour for the bottom of the cell</summary>
		private Colour bottomBorderColour;
		
		/// <summary> The background colour</summary>
		private Colour backgroundColour;
		/// <summary> The background pattern</summary>
		private Pattern pattern;
		/// <summary> The options mask which is used to store the processed cell options (such
		/// as alignment, borders etc)
		/// </summary>
		private int options;
		/// <summary> The index of this XF record within the workbook</summary>
		private int xfIndex;
		/// <summary> The font object for this XF record</summary>
		private FontRecord font;
		/// <summary> The format object for this XF record.  This is used when creating
		/// a writable record
		/// </summary>
		private DisplayFormat format;
		/// <summary> Flag to indicate whether this XF record has been initialized</summary>
		private bool initialized;
		
		/// <summary> Indicates whether this cell was constructed by an API or read
		/// from an existing Excel file
		/// </summary>
		private bool read;
		
		/// <summary> The excel format for this record. This is used to display the actual
		/// excel format string back to the user (eg. when generating certain
		/// types of XML document) as opposed to the java equivalent
		/// </summary>
		private NExcel.Format.Format excelFormat;
		
		/// <summary> Flag to indicate whether the format information has been initialized.
		/// This is false if the xf record has been read in, but true if it
		/// has been written
		/// </summary>
		private bool formatInfoInitialized;
		
		/// <summary> Flag to indicate whether this cell was copied.  If it was copied, then
		/// it can be set to uninitialized, allowing us to change certain format
		/// information
		/// </summary>
		private bool copied;
		
		/// <summary> A handle to the formatting records.  The purpose of this is
		/// to read the formatting information back, for the purposes of
		/// output eg. to some form of XML
		/// </summary>
		private FormattingRecords formattingRecords;
		
		/// <summary> The list of built in date format values</summary>
		private static readonly int[] dateFormats = new int[]{0xe, 0xf, 0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x2d, 0x2e, 0x2f};
		
		/// <summary> The list of lang-specific date format equivalents</summary>
		private static readonly DateTimeFormatInfo[] langDateFormats = new DateTimeFormatInfo[]{
				new DateTimeFormatInfo("MM/dd/yyyy"), 
				new DateTimeFormatInfo("d-MMM-yy"), 
				new DateTimeFormatInfo("d-MMM"), 
				new DateTimeFormatInfo("MMM-yy"), 
				new DateTimeFormatInfo("h:mm a"), 
				new DateTimeFormatInfo("h:mm:ss a"), 
				new DateTimeFormatInfo("H:mm"), 
				new DateTimeFormatInfo("H:mm:ss"), 
				new DateTimeFormatInfo("M/d/yy H:mm"), 
				new DateTimeFormatInfo("mm:ss"), 
				new DateTimeFormatInfo("H:mm:ss"), 
				new DateTimeFormatInfo("mm:ss.S")
		};
		
		/// <summary> The list of built in number format values</summary>
		private static int[] numberFormats = new int[]{0x1, 0x2, 0x3, 0x4, 0x5, 0x6, 0x7, 0x8, 0x9, 0xa, 0xb, 0x25, 0x26, 0x27, 0x28, 0x29, 0x2a, 0x2b, 0x2c, 0x30};
		
		/// <summary>  The list of lang-specific number format equivalents</summary>
		private static NumberFormatInfo[] langNumberFormats = new NumberFormatInfo[]
		{
			new NumberFormatInfo("0"), 
			new NumberFormatInfo("0.00"), 
			new NumberFormatInfo("#,##0"), 
			new NumberFormatInfo("#,##0.00"), 
			new NumberFormatInfo("$#,##0;($#,##0)"), 
			new NumberFormatInfo("$#,##0;($#,##0)"), 
			new NumberFormatInfo("$#,##0.00;($#,##0.00)"), 
			new NumberFormatInfo("$#,##0.00;($#,##0.00)"), 
			new NumberFormatInfo("0%"), 
			new NumberFormatInfo("0.00%"), 
			new NumberFormatInfo("0.00E00"), 
			new NumberFormatInfo("#,##0;(#,##0)"), 
			new NumberFormatInfo("#,##0;(#,##0)"), 
			new NumberFormatInfo("#,##0.00;(#,##0.00)"), 
			new NumberFormatInfo("#,##0.00;(#,##0.00)"), 
			new NumberFormatInfo("#,##0;(#,##0)"), 
			new NumberFormatInfo("$#,##0;($#,##0)"), 
			new NumberFormatInfo("#,##0.00;(#,##0.00)"), 
			new NumberFormatInfo("$#,##0.00;($#,##0.00)"), 
			new NumberFormatInfo("##0.0E0")
		};
		
		
		
		// Type to distinguish between biff7 and biff8
		public class BiffType
		{
		}
		
		
		public static readonly BiffType biff8 = new BiffType();
		public static readonly BiffType biff7 = new BiffType();
		
		/// <summary> The biff type</summary>
		private BiffType biffType;
		
		// Type to distinguish between cell and style records
		public class XFType
		{
		}
		protected internal static readonly XFType cell = new XFType();
		protected internal static readonly XFType style = new XFType();
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="bt">the biff type
		/// </param>
		public XFRecord(Record t, BiffType bt):base(t)
		{
			
			biffType = bt;
			
			sbyte[] data = getRecord().Data;
			
			fontIndex = IntegerHelper.getInt(data[0], data[1]);
			formatIndex = IntegerHelper.getInt(data[2], data[3]);
			date = false;
			number = false;
			
			// Compare against the date formats
			for (int i = 0; i < dateFormats.Length; i++)
			{
				if (formatIndex == dateFormats[i])
				{
					date = true;
					dateFormat = langDateFormats[i];
				}
			}
			
			// Compare against the number formats
			for (int i = 0; i < numberFormats.Length; i++)
			{
				if (formatIndex == numberFormats[i])
				{
					number = true;
					numberFormat = langNumberFormats[i];
				}
			}
			
			// Initialize the parent format and the type
			int cellAttributes = IntegerHelper.getInt(data[4], data[5]);
			parentFormat = (cellAttributes & 0xfff0) >> 4;
			
			int formatType = cellAttributes & 0x4;
			xfFormatType = formatType == 0?cell:style;
			locked = ((cellAttributes & 0x1) != 0);
			hidden = ((cellAttributes & 0x2) != 0);
			
			if (xfFormatType == cell && (parentFormat & 0xfff) == 0xfff)
			{
				// Something is screwy with the parent format - set to zero
				parentFormat = 0;
				logger.warn("Invalid parent format found - ignoring");
			}
			
			initialized = false;
			read = true;
			formatInfoInitialized = false;
			copied = false;
		}
		
		/// <summary> A constructor used when creating a writable record
		/// 
		/// </summary>
		/// <param name="fnt">the font
		/// </param>
		/// <param name="form">the format
		/// </param>
		public XFRecord(FontRecord fnt, DisplayFormat form):base(NExcel.Biff.Type.XF)
		{
			
			initialized = false;
			locked = true;
			hidden = false;
			align = Alignment.GENERAL;
			valign = VerticalAlignment.BOTTOM;
			orientation = Orientation.HORIZONTAL;
			wrap = false;
			leftBorder = BorderLineStyle.NONE;
			rightBorder = BorderLineStyle.NONE;
			topBorder = BorderLineStyle.NONE;
			bottomBorder = BorderLineStyle.NONE;
			leftBorderColour = Colour.PALETTE_BLACK;
			rightBorderColour = Colour.PALETTE_BLACK;
			topBorderColour = Colour.PALETTE_BLACK;
			bottomBorderColour = Colour.PALETTE_BLACK;
			pattern = Pattern.NONE;
			backgroundColour = Colour.DEFAULT_BACKGROUND;
			shrinkToFit = false;
			
			// This will be set by the initialize method and the subclass respectively
			parentFormat = 0;
			xfFormatType = null;
			
			font = fnt;
			format = form;
			biffType = biff8;
			read = false;
			copied = false;
			formatInfoInitialized = true;
			
			Assert.verify(font != null);
			Assert.verify(format != null);
		}
		
		/// <summary> Copy constructor.  Used for copying writable formats, typically
		/// when duplicating formats to handle merged cells
		/// 
		/// </summary>
		/// <param name="fmt">XFRecord
		/// </param>
		protected internal XFRecord(XFRecord fmt):base(NExcel.Biff.Type.XF)
		{
			
			initialized = false;
			locked = fmt.locked;
			hidden = fmt.hidden;
			align = fmt.align;
			valign = fmt.valign;
			orientation = fmt.orientation;
			wrap = fmt.wrap;
			leftBorder = fmt.leftBorder;
			rightBorder = fmt.rightBorder;
			topBorder = fmt.topBorder;
			bottomBorder = fmt.bottomBorder;
			leftBorderColour = fmt.leftBorderColour;
			rightBorderColour = fmt.rightBorderColour;
			topBorderColour = fmt.topBorderColour;
			bottomBorderColour = fmt.bottomBorderColour;
			pattern = fmt.pattern;
			xfFormatType = fmt.xfFormatType;
			shrinkToFit = fmt.shrinkToFit;
			parentFormat = fmt.parentFormat;
			backgroundColour = fmt.backgroundColour;
			
			// Shallow copy is sufficient for these purposes
			font = fmt.font;
			format = fmt.format;
			
			fontIndex = fmt.fontIndex;
			formatIndex = fmt.formatIndex;
			
			formatInfoInitialized = fmt.formatInfoInitialized;
			
			biffType = biff8;
			read = false;
			copied = true;
		}
		
		/// <summary> A public copy constructor which can be used for copy formats between
		/// different sheets.  Unlike the the other copy constructor, this
		/// version does a deep copy
		/// 
		/// </summary>
		/// <param name="cellFormat">the format to copy
		/// </param>
		protected internal XFRecord(CellFormat cellFormat):base(NExcel.Biff.Type.XF)
		{
			
			Assert.verify(cellFormat is XFRecord);
			XFRecord fmt = (XFRecord) cellFormat;
			
			if (!fmt.formatInfoInitialized)
			{
				fmt.initializeFormatInformation();
			}
			
			locked = fmt.locked;
			hidden = fmt.hidden;
			align = fmt.align;
			valign = fmt.valign;
			orientation = fmt.orientation;
			wrap = fmt.wrap;
			leftBorder = fmt.leftBorder;
			rightBorder = fmt.rightBorder;
			topBorder = fmt.topBorder;
			bottomBorder = fmt.bottomBorder;
			leftBorderColour = fmt.leftBorderColour;
			rightBorderColour = fmt.rightBorderColour;
			topBorderColour = fmt.topBorderColour;
			bottomBorderColour = fmt.bottomBorderColour;
			pattern = fmt.pattern;
			xfFormatType = fmt.xfFormatType;
			parentFormat = fmt.parentFormat;
			shrinkToFit = fmt.shrinkToFit;
			backgroundColour = fmt.backgroundColour;
			
			// Deep copy of the font
			font = new FontRecord(fmt.Font);
			
			// Copy the format
			if (fmt.Format == null)
			{
				// format is writable
				if (fmt.format.isBuiltIn())
				{
					format = fmt.format;
				}
				else
				{
					// Format is not built in, so do a deep copy
					format = new FormatRecord((FormatRecord) fmt.format);
				}
			}
			else if (fmt.Format is BuiltInFormat)
			{
				// read excel format is built in
				excelFormat = (BuiltInFormat) fmt.excelFormat;
				format = (BuiltInFormat) fmt.excelFormat;
			}
			else
			{
				// read excel format is user defined
				Assert.verify(fmt.formatInfoInitialized);
				
				// in this case FormattingRecords should initialize the excelFormat
				// field with an instance of FormatRecord
				Assert.verify(fmt.excelFormat is FormatRecord);
				
				// Format is not built in, so do a deep copy
				FormatRecord fr = new FormatRecord((FormatRecord) fmt.excelFormat);
				
				// Set both format fields to be the same object, since
				// FormatRecord implements all the necessary interfaces
				excelFormat = fr;
				format = fr;
			}
			
			biffType = biff8;
			
			// The format info should be all OK by virtue of the deep copy
			formatInfoInitialized = true;
			
			
			// This format was not read in
			read = false;
			
			// Treat this as a new cell record, so set the copied flag to false
			copied = false;
			
			// The font or format indexes need to be set, so set initialized to false
			initialized = false;
		}
		
		/// <summary> Converts the various fields into binary data.  If this object has
		/// been read from an Excel file rather than being requested by a user (ie.
		/// if the read flag is TRUE) then
		/// no processing takes place and the raw data is simply returned.
		/// 
		/// </summary>
		/// <returns> the raw data for writing
		/// </returns>
		public override sbyte[] getData()
		{
			// Format rationalization process means that we always want to
			// regenerate the format info - even if the spreadsheet was
			// read in
			if (!formatInfoInitialized)
			{
				initializeFormatInformation();
			}
			
			sbyte[] data = new sbyte[20];
			
			IntegerHelper.getTwoBytes(fontIndex, data, 0);
			IntegerHelper.getTwoBytes(formatIndex, data, 2);
			
			// Do the cell attributes
			int cellAttributes = 0;
			
			if (this.Locked)
			{
				cellAttributes |= 0x01;
			}
			
			if (Hidden)
			{
				cellAttributes |= 0x02;
			}
			
			if (xfFormatType == style)
			{
				cellAttributes |= 0x04;
				parentFormat = 0xffff;
			}
			
			cellAttributes |= (parentFormat << 4);
			
			IntegerHelper.getTwoBytes(cellAttributes, data, 4);
			
			int alignMask = align.Value;
			
			if (wrap)
			{
				alignMask |= 0x08;
			}
			
			alignMask |= (valign.Value << 4);
			
			alignMask |= (orientation.Value << 8);
			
			IntegerHelper.getTwoBytes(alignMask, data, 6);
			
			// Set the borders
			int borderMask = leftBorder.Value;
			borderMask |= (rightBorder.Value << 4);
			borderMask |= (topBorder.Value << 8);
			borderMask |= (bottomBorder.Value << 12);
			
			IntegerHelper.getTwoBytes(borderMask, data, 10);
			
			// Set the border palette information if border mask is non zero
			// Hard code the colours to be black
			if (borderMask != 0)
			{
				sbyte lc = (sbyte) leftBorderColour.Value;
				sbyte rc = (sbyte) rightBorderColour.Value;
				sbyte tc = (sbyte) topBorderColour.Value;
				sbyte bc = (sbyte) bottomBorderColour.Value;
				
				data[12] = (sbyte) ((lc & 0x7f) | ((rc & 0x01) << 7));
				data[13] = (sbyte) ((rc & 0x7f) >> 1);
				data[14] = (sbyte) ((tc & 0x7f) | ((bc & 0x01) << 7));
				data[15] = (sbyte) ((bc & 0x7f) >> 1);
			}
			
			// Set the background pattern
			IntegerHelper.getTwoBytes(pattern.Value, data, 16);
			
			// Set the colour palette
			int colourPaletteMask = backgroundColour.Value;
			colourPaletteMask |= (0x40 << 7);
			IntegerHelper.getTwoBytes(colourPaletteMask, data, 18);
			
			// Set the cell options
			if (shrinkToFit)
			{
				options |= 0x10;
			}
			else
			{
				options &= 0xffef;
			}
			
			IntegerHelper.getTwoBytes(options, data, 8);
			
			return data;
		}
		
		/// <summary> Accessor for the locked flag
		/// 
		/// </summary>
		/// <returns> TRUE if this XF record locks cells, FALSE otherwise
		/// </returns>
		public bool Locked
		{
		get
		{
		return locked;
		}
		}
		
		/// <summary> Accessor for whether a particular cell is locked
		/// 
		/// </summary>
		/// <returns> TRUE if this cell is locked, FALSE otherwise
		/// </returns>
		public virtual bool isLocked()
		{
			if (!formatInfoInitialized)
			{
				initializeFormatInformation();
			}
			
			return locked;
		}
		
		/// <summary> Sets the horizontal alignment for the data in this cell.
		/// This method should only be called from its writable subclass
		/// CellXFRecord
		/// 
		/// </summary>
		/// <param name="c">the background colour
		/// </param>
		/// <param name="p">the background pattern
		/// </param>
		protected internal virtual void  setXFBackground(Colour c, Pattern p)
		{
			Assert.verify(!initialized);
			backgroundColour = c;
			pattern = p;
		}
		
		/// <summary> Sets the border for this cell
		/// This method should only be called from its writable subclass
		/// CellXFRecord
		/// 
		/// </summary>
		/// <param name="b">the border
		/// </param>
		/// <param name="ls">the border line style
		/// </param>
		protected internal virtual void  setXFBorder(Border b, BorderLineStyle ls, Colour c)
		{
			Assert.verify(!initialized);
			
			if (c == Colour.BLACK)
			{
				c = Colour.PALETTE_BLACK;
			}
			
			if (b == Border.LEFT)
			{
				leftBorder = ls;
				leftBorderColour = c;
			}
			else if (b == Border.RIGHT)
			{
				rightBorder = ls;
				rightBorderColour = c;
			}
			else if (b == Border.TOP)
			{
				topBorder = ls;
				topBorderColour = c;
			}
			else if (b == Border.BOTTOM)
			{
				bottomBorder = ls;
				bottomBorderColour = c;
			}
			return ;
		}
		
		
		/// <summary> Gets the line style for the given cell border
		/// If a border type of ALL or NONE is specified, then a line style of
		/// NONE is returned
		/// 
		/// </summary>
		/// <param name="border">the cell border we are interested in
		/// </param>
		/// <returns> the line style of the specified border
		/// </returns>
		public virtual BorderLineStyle getBorder(Border border)
		{
			return getBorderLine(border);
		}
		
		/// <summary> Gets the line style for the given cell border
		/// If a border type of ALL or NONE is specified, then a line style of
		/// NONE is returned
		/// 
		/// </summary>
		/// <param name="border">the cell border we are interested in
		/// </param>
		/// <returns> the line style of the specified border
		/// </returns>
		public virtual BorderLineStyle getBorderLine(Border border)
		{
			// Don't bother with the short cut records
			if (border == Border.NONE || border == Border.ALL)
			{
				return BorderLineStyle.NONE;
			}
			
			if (!formatInfoInitialized)
			{
				initializeFormatInformation();
			}
			
			if (border == Border.LEFT)
			{
				return leftBorder;
			}
			else if (border == Border.RIGHT)
			{
				return rightBorder;
			}
			else if (border == Border.TOP)
			{
				return topBorder;
			}
			else if (border == Border.BOTTOM)
			{
				return bottomBorder;
			}
			
			return BorderLineStyle.NONE;
		}
		
		/// <summary> Gets the line style for the given cell border
		/// If a border type of ALL or NONE is specified, then a line style of
		/// NONE is returned
		/// 
		/// </summary>
		/// <param name="border">the cell border we are interested in
		/// </param>
		/// <returns> the line style of the specified border
		/// </returns>
		public virtual Colour getBorderColour(Border border)
		{
			// Don't bother with the short cut records
			if (border == Border.NONE || border == Border.ALL)
			{
				return Colour.PALETTE_BLACK;
			}
			
			if (!formatInfoInitialized)
			{
				initializeFormatInformation();
			}
			
			if (border == Border.LEFT)
			{
				return leftBorderColour;
			}
			else if (border == Border.RIGHT)
			{
				return rightBorderColour;
			}
			else if (border == Border.TOP)
			{
				return topBorderColour;
			}
			else if (border == Border.BOTTOM)
			{
				return bottomBorderColour;
			}
			
			return Colour.BLACK;
		}
		
		
		/// <summary> Determines if this cell format has any borders at all.  Used to
		/// set the new borders when merging a group of cells
		/// 
		/// </summary>
		/// <returns> TRUE if this cell has any borders, FALSE otherwise
		/// </returns>
		public bool hasBorders()
		{
			if (!formatInfoInitialized)
			{
				initializeFormatInformation();
			}
			
			if (leftBorder == BorderLineStyle.NONE && rightBorder == BorderLineStyle.NONE && topBorder == BorderLineStyle.NONE && bottomBorder == BorderLineStyle.NONE)
			{
				return false;
			}
			
			return true;
		}
		
		/// <summary> If this cell has not been read in from an existing Excel sheet,
		/// then initializes this record with the XF index passed in. Calls
		/// initialized on the font and format record
		/// 
		/// </summary>
		/// <param name="pos">the xf index to initialize this record with
		/// </param>
		/// <param name="fr">the containing formatting records
		/// </param>
		/// <param name="fonts">the container for the fonts
		/// </param>
		/// <exception cref=""> NumFormatRecordsException
		/// </exception>
		public void  initialize(int pos, FormattingRecords fr, Fonts fonts)
		{
			xfIndex = pos;
			formattingRecords = fr;
			
			// If this file has been read in or copied,
			// the font and format indexes will
			// already be initialized, so just set the initialized flag and
			// return
			if (read || copied)
			{
				initialized = true;
				return ;
			}
			
			if (!font.IsInitialized())
			{
				fonts.addFont(font);
			}
			
			if (!format.isInitialized())
			{
				fr.addFormat(format);
			}
			
			fontIndex = font.FontIndex;
			formatIndex = format.FormatIndex;
			
			initialized = true;
		}
		
		/// <summary> Resets the initialize flag.  This is called by the constructor of
		/// WritableWorkbookImpl to reset the statically declared fonts
		/// </summary>
		public void  uninitialize()
		{
			initialized = false;
		}
		
		/// <summary> Sets the XF index.  Called when rationalizing the XF records
		/// immediately prior to writing
		/// 
		/// </summary>
		/// <param name="xfi">the new xf index
		/// </param>
		internal void  setXFIndex(int xfi)
		{
			xfIndex = xfi;
		}
		
		/// <summary> Accessor for the XF index
		/// 
		/// </summary>
		/// <returns> the XF index for this cell
		/// </returns>
		public int getXFIndex()
		{
			return xfIndex;
		}
		
		/// <summary> Gets the font used by this format
		/// 
		/// </summary>
		/// <returns> the font
		/// </returns>
		public virtual Font Font
		{
		get
		{
		if (!formatInfoInitialized)
		{
		initializeFormatInformation();
		}
		return font;
		}
		}
		
		
		/// <summary> Initializes the internal format information from the data read in</summary>
		private void  initializeFormatInformation()
		{
			// Initialize the cell format string
			if (formatIndex < BuiltInFormat.builtIns.Length && BuiltInFormat.builtIns[formatIndex] != null)
			{
				excelFormat = BuiltInFormat.builtIns[formatIndex];
			}
			else
			{
				excelFormat = formattingRecords.getFormatRecord(formatIndex);
			}
			
			// Initialize the font
			font = formattingRecords.Fonts.getFont(fontIndex);
			
			// Initialize the cell format data from the binary record
			sbyte[] data = getRecord().Data;
			
			// Get the parent record
			int cellAttributes = IntegerHelper.getInt(data[4], data[5]);
			parentFormat = (cellAttributes & 0xfff0) >> 4;
			int formatType = cellAttributes & 0x4;
			xfFormatType = formatType == 0?cell:style;
			locked = ((cellAttributes & 0x1) != 0);
			hidden = ((cellAttributes & 0x2) != 0);
			
			if (xfFormatType == cell && (parentFormat & 0xfff) == 0xfff)
			{
				// Something is screwy with the parent format - set to zero
				parentFormat = 0;
				logger.warn("Invalid parent format found - ignoring");
			}
			
			
			int alignMask = IntegerHelper.getInt(data[6], data[7]);
			
			// Get the wrap
			if ((alignMask & 0x08) != 0)
			{
				wrap = true;
			}
			
			// Get the horizontal alignment
			align = Alignment.getAlignment(alignMask & 0x7);
			
			// Get the vertical alignment
			valign = VerticalAlignment.getAlignment((alignMask >> 4) & 0x7);
			
			// Get the orientation
			orientation = Orientation.getOrientation((alignMask >> 8) & 0xff);
			
			int attr = IntegerHelper.getInt(data[8], data[9]);
			
			// Get the shrink to fit flag
			shrinkToFit = (attr & 0x10) != 0;
			
			// Get the used attribute
			if (biffType == biff8)
			{
				usedAttributes = data[9];
			}
			
			// Get the borders
			int borderMask = IntegerHelper.getInt(data[10], data[11]);
			
			leftBorder = BorderLineStyle.getStyle(borderMask & 0x7);
			rightBorder = BorderLineStyle.getStyle((borderMask >> 4) & 0x7);
			topBorder = BorderLineStyle.getStyle((borderMask >> 8) & 0x7);
			bottomBorder = BorderLineStyle.getStyle((borderMask >> 12) & 0x7);
			
			int borderColourMask = IntegerHelper.getInt(data[12], data[13]);
			
			leftBorderColour = Colour.getInternalColour(borderColourMask & 0x7f);
			rightBorderColour = Colour.getInternalColour((borderColourMask & 0x3f80) >> 7);
			
			borderColourMask = IntegerHelper.getInt(data[14], data[15]);
			topBorderColour = Colour.getInternalColour(borderColourMask & 0x7f);
			bottomBorderColour = Colour.getInternalColour((borderColourMask & 0x3f80) >> 7);
			
			if (biffType == biff8)
			{
				// Get the background pattern
				int patternVal = IntegerHelper.getInt(data[16], data[17]);
				pattern = Pattern.getPattern(patternVal);
				
				// Get the background colour
				int colourPaletteMask = IntegerHelper.getInt(data[18], data[19]);
				backgroundColour = Colour.getInternalColour(colourPaletteMask & 0x3f);
				
				
				if (backgroundColour == Colour.UNKNOWN || backgroundColour == Colour.DEFAULT_BACKGROUND1)
				{
					backgroundColour = Colour.DEFAULT_BACKGROUND;
				}
			}
			else
			{
				pattern = Pattern.NONE;
				backgroundColour = Colour.DEFAULT_BACKGROUND;
			}
			
			// Set the lazy initialization flag
			formatInfoInitialized = true;
		}
		
		/// <summary> Standard hash code implementation</summary>
		/// <returns> the hash code
		/// </returns>
		public override int GetHashCode()
		{
			return 17;
		}
		
		/// <summary> Equals method.  This is called when comparing writable formats
		/// in order to prevent duplicate formats being added to the workbook
		/// 
		/// </summary>
		/// <param name="o">object to compare
		/// </param>
		/// <returns> TRUE if the objects are equal, FALSE otherwise
		/// </returns>
		public  override bool Equals(System.Object o)
		{
			if (o == this)
			{
				return true;
			}
			
			if (!(o is XFRecord))
			{
				return false;
			}
			
			XFRecord xfr = (XFRecord) o;
			
			// Both records must be writable and have their format info initialized
			if (!formatInfoInitialized)
			{
				initializeFormatInformation();
			}
			
			if (!xfr.formatInfoInitialized)
			{
				xfr.initializeFormatInformation();
			}
			
			if (xfFormatType != xfr.xfFormatType || parentFormat != xfr.parentFormat || locked != xfr.locked || hidden != xfr.hidden || usedAttributes != xfr.usedAttributes)
			{
				return false;
			}
			
			if (align != xfr.align || valign != xfr.valign || orientation != xfr.orientation || wrap != xfr.wrap || shrinkToFit != xfr.shrinkToFit)
			{
				return false;
			}
			
			if (leftBorder != xfr.leftBorder || rightBorder != xfr.rightBorder || topBorder != xfr.topBorder || bottomBorder != xfr.bottomBorder)
			{
				return false;
			}
			
			if (leftBorderColour != xfr.leftBorderColour || rightBorderColour != xfr.rightBorderColour || topBorderColour != xfr.topBorderColour || bottomBorderColour != xfr.bottomBorderColour)
			{
				return false;
			}
			
			if (backgroundColour != xfr.backgroundColour || pattern != xfr.pattern)
			{
				return false;
			}
			
			// Sufficient to just do shallow equals on font, format objects,
			// since we are testing for the presence of clones anwyay
			// Use indices rather than objects because of the rationalization
			// process (which does not set the object on an XFRecord)
			if (fontIndex != xfr.fontIndex || formatIndex != xfr.formatIndex)
			{
				return false;
			}
			
			return true;
		}
		
		/// <summary> Sets the format type and parent format from the writable subclass</summary>
		/// <param name="t">the xf type
		/// </param>
		/// <param name="pf">the parent format
		/// </param>
		protected internal virtual void  setXFDetails(XFType t, int pf)
		{
			xfFormatType = t;
			parentFormat = pf;
		}
		
		/// <summary> Changes the appropriate indexes during the rationalization process</summary>
		/// <param name="xfMapping">the xf index re-mappings
		/// </param>
		internal virtual void  rationalize(IndexMapping xfMapping)
		{
			xfIndex = xfMapping.getNewIndex(xfIndex);
			
			if (xfFormatType == cell)
			{
				parentFormat = xfMapping.getNewIndex(parentFormat);
			}
		}
		
		/// <summary> Sets the font object with a workbook specific clone.  Called from 
		/// the CellValue object when the font has been identified as a statically
		/// shared font
		/// </summary>
		public virtual void  setFont(FontRecord f)
		{
			// This style cannot be initialized, otherwise it would mean it would
			// have been initialized with shared font
			Assert.verify(!initialized);
			
			font = f;
		}
		static XFRecord()
		{
			logger = Logger.getLogger(typeof(XFRecord));
		}
	}
}
