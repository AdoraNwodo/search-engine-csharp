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
using Assert = common.Assert;
using PageOrientation = NExcel.Format.PageOrientation;
using PaperSize = NExcel.Format.PaperSize;
namespace NExcel
{
	
	/// <summary> This is a bean which client applications may use to get/set various
	/// properties which are associated with a particular worksheet, such
	/// as headers and footers, page orientation etc.
	/// </summary>
	public sealed class SheetSettings
	{
		/// <summary> Accessor for the orientation
		/// 
		/// </summary>
		/// <returns> the orientation
		/// </returns>
		/// <summary> Sets the paper orientation for printing this sheet
		/// 
		/// </summary>
		/// <param name="po">the orientation
		/// </param>
		public PageOrientation Orientation
		{
			get
			{
				return orientation;
			}
			
			set
			{
				orientation = value;
			}
			
		}
		/// <summary> Accessor for the paper size
		/// 
		/// </summary>
		/// <returns> the paper size
		/// </returns>
		/// <summary> Sets the paper size to be used when printing this sheet
		/// 
		/// </summary>
		/// <param name="ps">the paper size
		/// </param>
		public PaperSize PaperSize
		{
			get
			{
				return paperSize;
			}
			
			set
			{
				paperSize = value;
			}
			
		}
		/// <summary> Queries whether this sheet is protected (ie. read only)
		/// 
		/// </summary>
		/// <returns> TRUE if this sheet is read only, FALSE otherwise
		/// </returns>
		/// <summary> Sets the protected (ie. read only) status of this sheet
		/// 
		/// </summary>
		/// <param name="p">the protected status
		/// </param>
		public bool Protected
		{
			get
			{
				return sheetProtected;
			}
			
			set
			{
				sheetProtected = value;
			}
			
		}
		/// <summary> Accessor for the header margin
		/// 
		/// </summary>
		/// <returns> the header margin
		/// </returns>
		/// <summary> Sets the margin for any page headers
		/// 
		/// </summary>
		/// <param name="d">the margin in inches
		/// </param>
		public double HeaderMargin
		{
			get
			{
				return headerMargin;
			}
			
			set
			{
				headerMargin = value;
			}
			
		}
		/// <summary> Accessor for the footer margin
		/// 
		/// </summary>
		/// <returns> the footer margin
		/// </returns>
		/// <summary> Sets the margin for any page footer
		/// 
		/// </summary>
		/// <param name="d">the footer margin in inches
		/// </param>
		public double FooterMargin
		{
			get
			{
				return footerMargin;
			}
			
			set
			{
				footerMargin = value;
			}
			
		}
		/// <summary> Accessor for the hidden nature of this sheet
		/// 
		/// </summary>
		/// <returns> TRUE if this sheet is hidden, FALSE otherwise
		/// </returns>
		/// <summary> Sets the hidden status of this worksheet
		/// 
		/// </summary>
		/// <param name="h">the hidden flag
		/// </param>
		public bool Hidden
		{
			get
			{
				return hidden;
			}
			
			set
			{
				hidden = value;
			}
			
		}
		/// <summary> Accessor for the scale factor
		/// 
		/// </summary>
		/// <returns> the scale factor
		/// </returns>
		/// <summary> Sets the scale factor for this sheet to be used when printing.  The
		/// parameter is a percentage, therefore setting a scale factor of 100 will
		/// print at normal size, 50 half size, 200 double size etc
		/// 
		/// </summary>
		/// <param name="sf">the scale factor as a percentage
		/// </param>
		public int ScaleFactor
		{
			get
			{
				return scaleFactor;
			}
			
			set
			{
				scaleFactor = value;
				fitToPages = false;
			}
			
		}
		/// <summary> Accessor for the page start
		/// 
		/// </summary>
		/// <returns> the page start
		/// </returns>
		/// <summary> Sets the page number at which to commence printing
		/// 
		/// </summary>
		/// <param name="ps">the page start number
		/// </param>
		public int PageStart
		{
			get
			{
				return pageStart;
			}
			
			set
			{
				pageStart = value;
			}
			
		}
		/// <summary> Accessor for the fit width
		/// 
		/// </summary>
		/// <returns> the number of pages this sheet will be printed into widthwise
		/// </returns>
		/// <summary> Sets the number of pages widthwise which this sheet should be
		/// printed into
		/// 
		/// </summary>
		/// <param name="fw">the number of pages
		/// </param>
		public int FitWidth
		{
			get
			{
				return fitWidth;
			}
			
			set
			{
				fitWidth = value;
				fitToPages = true;
			}
			
		}
		/// <summary> Accessor for the fit height
		/// 
		/// </summary>
		/// <returns> the number of pages this sheet will be printed into heightwise
		/// </returns>
		/// <summary> Sets the number of pages vertically that this sheet will be printed into
		/// 
		/// </summary>
		/// <param name="fh">the number of pages this sheet will be printed into heightwise
		/// </param>
		public int FitHeight
		{
			get
			{
				return fitHeight;
			}
			
			set
			{
				fitHeight = value;
				fitToPages = true;
			}
			
		}
		/// <summary> Accessor for the horizontal print resolution
		/// 
		/// </summary>
		/// <returns> the horizontal print resolution
		/// </returns>
		/// <summary> Sets the horizontal print resolution
		/// 
		/// </summary>
		/// <param name="hpw">the print resolution
		/// </param>
		public int HorizontalPrintResolution
		{
			get
			{
				return horizontalPrintResolution;
			}
			
			set
			{
				horizontalPrintResolution = value;
			}
			
		}
		/// <summary> Accessor for the vertical print resolution
		/// 
		/// </summary>
		/// <returns> the vertical print resolution
		/// </returns>
		/// <summary> Sets the vertical print reslution
		/// 
		/// </summary>
		/// <param name="vpw">the vertical print resolution
		/// </param>
		public int VerticalPrintResolution
		{
			get
			{
				return verticalPrintResolution;
			}
			
			set
			{
				verticalPrintResolution = value;
			}
			
		}
		/// <summary> Accessor for the right margin
		/// 
		/// </summary>
		/// <returns> the right margin in inches
		/// </returns>
		/// <summary> Sets the right margin
		/// 
		/// </summary>
		/// <param name="m">the right margin in inches
		/// </param>
		public double RightMargin
		{
			get
			{
				return rightMargin;
			}
			
			set
			{
				rightMargin = value;
			}
			
		}
		/// <summary> Accessor for the left margin
		/// 
		/// </summary>
		/// <returns> the left margin in inches
		/// </returns>
		/// <summary> Sets the left margin
		/// 
		/// </summary>
		/// <param name="m">the left margin in inches
		/// </param>
		public double LeftMargin
		{
			get
			{
				return leftMargin;
			}
			
			set
			{
				leftMargin = value;
			}
			
		}
		/// <summary> Accessor for the top margin
		/// 
		/// </summary>
		/// <returns> the top margin in inches
		/// </returns>
		/// <summary> Sets the top margin
		/// 
		/// </summary>
		/// <param name="m">the top margin in inches
		/// </param>
		public double TopMargin
		{
			get
			{
				return topMargin;
			}
			
			set
			{
				topMargin = value;
			}
			
		}
		/// <summary> Accessor for the bottom margin
		/// 
		/// </summary>
		/// <returns> the bottom margin in inches
		/// </returns>
		/// <summary> Sets the bottom margin
		/// 
		/// </summary>
		/// <param name="m">the bottom margin in inches
		/// </param>
		public double BottomMargin
		{
			get
			{
				return bottomMargin;
			}
			
			set
			{
				bottomMargin = value;
			}
			
		}
		/// <summary> Gets the default margin width
		/// 
		/// </summary>
		/// <returns> the default margin width
		/// </returns>
		public double DefaultWidthMargin
		{
			get
			{
				return defaultWidthMargin;
			}
			
		}
		/// <summary> Gets the default margin height
		/// 
		/// </summary>
		/// <returns> the default margin height
		/// </returns>
		public double DefaultHeightMargin
		{
			get
			{
				return defaultHeightMargin;
			}
			
		}
		/// <summary> Accessor for the fit width print flag</summary>
		/// <returns> TRUE if the print is to fit to pages, false otherwise
		/// </returns>
		/// <summary> Accessor for the fit to pages flag</summary>
		/// <param name="b">TRUE to fit to pages, FALSE to use a scale factor
		/// </param>
		public bool FitToPages
		{
			get
			{
				return fitToPages;
			}
			
			set
			{
				fitToPages = value;
			}
			
		}
		/// <summary> Accessor for the password
		/// 
		/// </summary>
		/// <returns> the password to unlock this sheet, or NULL if not protected
		/// </returns>
		/// <summary> Sets the password for this sheet
		/// 
		/// </summary>
		/// <param name="s">the password
		/// </param>
		public string Password
		{
			get
			{
				return password;
			}
			
			set
			{
				password = value;
			}
			
		}
		/// <summary> Accessor for the password hash - used only when copying sheets
		/// 
		/// </summary>
		/// <returns> passwordHash
		/// </returns>
		/// <summary> Accessor for the password hash - used only when copying sheets
		/// 
		/// </summary>
		/// <param name="ph">the password hash
		/// </param>
		public int PasswordHash
		{
			get
			{
				return passwordHash;
			}
			
			set
			{
				passwordHash = value;
			}
			
		}
		/// <summary> Accessor for the default column width
		/// 
		/// </summary>
		/// <returns> the default column width, in characters
		/// </returns>
		/// <summary> Sets the default column width
		/// 
		/// </summary>
		/// <param name="w">the new default column width
		/// </param>
		public int DefaultColumnWidth
		{
			get
			{
				return defaultColumnWidth;
			}
			
			set
			{
				defaultColumnWidth = value;
			}
			
		}
		/// <summary> Accessor for the default row height
		/// 
		/// </summary>
		/// <returns> the default row height, in 1/20ths of a point
		/// </returns>
		/// <summary> Sets the default row height
		/// 
		/// </summary>
		/// <param name="h">the default row height, in 1/20ths of a point
		/// </param>
		public int DefaultRowHeight
		{
			get
			{
				return defaultRowHeight;
			}
			
			set
			{
				defaultRowHeight = value;
			}
			
		}
		/// <summary> Accessor for the zoom factor.  Do not confuse zoom factor (which relates
		/// to the on screen view) with scale factor (which refers to the scale factor
		/// when printing)
		/// 
		/// </summary>
		/// <returns> the zoom factor as a percentage
		/// </returns>
		/// <summary> Sets the zoom factor.  Do not confuse zoom factor (which relates
		/// to the on screen view) with scale factor (which refers to the scale factor
		/// when printing)
		/// 
		/// </summary>
		/// <param name="zf">the zoom factor as a percentage
		/// </param>
		public int ZoomFactor
		{
			get
			{
				return zoomFactor;
			}
			
			set
			{
				zoomFactor = value;
			}
			
		}
		/// <summary> Accessor for the displayZeroValues property
		/// 
		/// </summary>
		/// <returns> TRUE to display zero values, FALSE not to bother
		/// </returns>
		/// <summary> Sets the displayZeroValues property
		/// 
		/// </summary>
		/// <param name="b">TRUE to show zero values, FALSE not to bother
		/// </param>
		public bool DisplayZeroValues
		{
			get
			{
				return displayZeroValues;
			}
			
			set
			{
				displayZeroValues = value;
			}
			
		}
		/// <summary> Accessor for the showGridLines property
		/// 
		/// </summary>
		/// <returns> TRUE if grid lines will be shown, FALSE otherwise
		/// </returns>
		/// <summary> Sets the showGridLines property
		/// 
		/// </summary>
		/// <param name="b">TRUE to show grid lines on this sheet, FALSE otherwise
		/// </param>
		public bool ShowGridLines
		{
			get
			{
				return showGridLines;
			}
			
			set
			{
				showGridLines = value;
			}
			
		}
		/// <summary> Accessor for the printGridLines property
		/// 
		/// </summary>
		/// <returns> TRUE if grid lines will be printed, FALSE otherwise
		/// </returns>
		/// <summary> Sets the printGridLines property
		/// 
		/// </summary>
		/// <param name="b">TRUE to print grid lines on this sheet, FALSE otherwise
		/// </param>
		public bool PrintGridLines
		{
			get
			{
				return printGridLines;
			}
			
			set
			{
				printGridLines = value;
			}
			
		}
		/// <summary> Accessor for the printHeaders property
		/// 
		/// </summary>
		/// <returns> TRUE if headers will be printed, FALSE otherwise
		/// </returns>
		/// <summary> Sets the printHeaders property
		/// 
		/// </summary>
		/// <param name="b">TRUE to print headers on this sheet, FALSE otherwise
		/// </param>
		public bool PrintHeaders
		{
			get
			{
				return printHeaders;
			}
			
			set
			{
				printHeaders = value;
			}
			
		}
		/// <summary> Gets the row at which the pane is frozen horizontally
		/// 
		/// </summary>
		/// <returns> the row at which the pane is horizontally frozen, or 0 if there
		/// is no freeze
		/// </returns>
		/// <summary> Sets the row at which the pane is frozen horizontally
		/// 
		/// </summary>
		/// <param name="row">the row number to freeze at
		/// </param>
		public int HorizontalFreeze
		{
			get
			{
				return horizontalFreeze;
			}
			
			set
			{
				horizontalFreeze = System.Math.Max(value, 0);
			}
			
		}
		/// <summary> Gets the column at which the pane is frozen vertically
		/// 
		/// </summary>
		/// <returns> the column at which the pane is vertically frozen, or 0 if there
		/// is no freeze
		/// </returns>
		/// <summary> Sets the row at which the pane is frozen vertically
		/// 
		/// </summary>
		/// <param name="col">the column number to freeze at
		/// </param>
		public int VerticalFreeze
		{
			get
			{
				return verticalFreeze;
			}
			
			set
			{
				verticalFreeze = System.Math.Max(value, 0);
			}
			
		}
		/// <summary> Accessor for the number of copies to print
		/// 
		/// </summary>
		/// <returns> the number of copies
		/// </returns>
		/// <summary> Sets the number of copies
		/// 
		/// </summary>
		/// <param name="c">the number of copies
		/// </param>
		public int Copies
		{
			get
			{
				return copies;
			}
			
			set
			{
				copies = value;
			}
			
		}
		/// <summary> Accessor for the header
		/// 
		/// </summary>
		/// <returns> the header
		/// </returns>
		/// <summary> Sets the header
		/// 
		/// </summary>
		/// <param name="h">the header
		/// </param>
		public HeaderFooter Header
		{
			get
			{
				return header;
			}
			
			set
			{
				header = value;
			}
			
		}
		/// <summary> Accessor for the footer
		/// 
		/// </summary>
		/// <returns> the footer
		/// </returns>
		/// <summary> Sets the footer
		/// 
		/// </summary>
		/// <param name="f">the footer
		/// </param>
		public HeaderFooter Footer
		{
			get
			{
				return footer;
			}
			
			set
			{
				footer = value;
			}
			
		}
		/// <summary> The page orientation</summary>
		private PageOrientation orientation;
		
		/// <summary> The paper size for printing</summary>
		private PaperSize paperSize;
		
		/// <summary> Indicates whether or not this sheet is protected</summary>
		private bool sheetProtected;
		
		/// <summary> Indicates whether or not this sheet is hidden</summary>
		private bool hidden;
		
		/// <summary> Indicates whether or not this sheet is selected</summary>
		private bool selected;
		
		/// <summary> The header</summary>
		private HeaderFooter header;
		
		/// <summary> The margin allocated for any page headers, in inches</summary>
		private double headerMargin;
		
		/// <summary> The footer</summary>
		private HeaderFooter footer;
		
		/// <summary> The margin allocated for any page footers, in inches</summary>
		private double footerMargin;
		
		/// <summary> The scale factor used when printing</summary>
		private int scaleFactor;
		
		/// <summary> The zoom factor used when viewing.  Note the difference between
		/// this and the scaleFactor which is used when printing
		/// </summary>
		private int zoomFactor;
		
		/// <summary> The page number at which to commence printing</summary>
		private int pageStart;
		
		/// <summary> The number of pages into which this excel sheet is squeezed widthwise</summary>
		private int fitWidth;
		
		/// <summary> The number of pages into which this excel sheet is squeezed heightwise</summary>
		private int fitHeight;
		
		/// <summary> The horizontal print resolution</summary>
		private int horizontalPrintResolution;
		
		/// <summary> The vertical print resolution</summary>
		private int verticalPrintResolution;
		
		/// <summary> The margin from the left hand side of the paper in inches</summary>
		private double leftMargin;
		
		/// <summary> The margin from the right hand side of the paper in inches</summary>
		private double rightMargin;
		
		/// <summary> The margin from the top of the paper in inches</summary>
		private double topMargin;
		
		/// <summary> The margin from the bottom of the paper in inches</summary>
		private double bottomMargin;
		
		/// <summary> Indicates whether to fit the print to the pages or scale the output
		/// This field is manipulated indirectly by virtue of the setFitWidth/Height
		/// methods
		/// </summary>
		private bool fitToPages;
		
		/// <summary> Indicates whether grid lines should be displayed</summary>
		private bool showGridLines;
		
		/// <summary> Indicates whether grid lines should be printed</summary>
		private bool printGridLines;
		
		/// <summary> Indicates whether sheet headings should be printed</summary>
		private bool printHeaders;
		
		/// <summary> Indicates whether the sheet should display zero values</summary>
		private bool displayZeroValues;
		
		/// <summary> The password for protected sheets</summary>
		private string password;
		
		/// <summary> The password hashcode - used when copying sheets</summary>
		private int passwordHash;
		
		/// <summary> The default column width, in characters</summary>
		private int defaultColumnWidth;
		
		/// <summary> The default row height, in 1/20th of a point</summary>
		private int defaultRowHeight;
		
		/// <summary> The horizontal freeze pane</summary>
		private int horizontalFreeze;
		
		/// <summary> The vertical freeze position</summary>
		private int verticalFreeze;
		
		/// <summary> The number of copies to print</summary>
		private int copies;
		
		// ***
		// The defaults
		// **
		private static readonly PageOrientation defaultOrientation = PageOrientation.PORTRAIT;
		private static readonly PaperSize defaultPaperSize = PaperSize.A4;
		private const double defaultHeaderMargin = 0.5;
		private const double defaultFooterMargin = 0.5;
		private const int defaultPrintResolution = 0x12c;
		private const double defaultWidthMargin = 0.75;
		private const double defaultHeightMargin = 1;
		
		private const int defaultDefaultColumnWidth = 8;
		private const int defaultDefaultRowHeight = 0xff;
		private const int defaultZoomFactor = 100;
		
		/// <summary> Default constructor</summary>
		public SheetSettings()
		{
			orientation = defaultOrientation;
			paperSize = defaultPaperSize;
			sheetProtected = false;
			hidden = false;
			selected = false;
			headerMargin = defaultHeaderMargin;
			footerMargin = defaultFooterMargin;
			horizontalPrintResolution = defaultPrintResolution;
			verticalPrintResolution = defaultPrintResolution;
			leftMargin = defaultWidthMargin;
			rightMargin = defaultWidthMargin;
			topMargin = defaultHeightMargin;
			bottomMargin = defaultHeightMargin;
			fitToPages = false;
			showGridLines = true;
			printGridLines = false;
			printHeaders = false;
			displayZeroValues = true;
			defaultColumnWidth = defaultDefaultColumnWidth;
			defaultRowHeight = defaultDefaultRowHeight;
			zoomFactor = defaultZoomFactor;
			horizontalFreeze = 0;
			verticalFreeze = 0;
			copies = 1;
			header = new HeaderFooter();
			footer = new HeaderFooter();
		}
		
		/// <summary> Copy constructor.  Called when copying sheets</summary>
		/// <param name="copy">the settings to copy
		/// </param>
		public SheetSettings(SheetSettings copy)
		{
			Assert.verify(copy != null);
			
			orientation = copy.orientation;
			paperSize = copy.paperSize;
			sheetProtected = copy.sheetProtected;
			hidden = copy.hidden;
			selected = false; // don't copy the selected flag
			headerMargin = copy.headerMargin;
			footerMargin = copy.footerMargin;
			scaleFactor = copy.scaleFactor;
			pageStart = copy.pageStart;
			fitWidth = copy.fitWidth;
			fitHeight = copy.fitHeight;
			horizontalPrintResolution = copy.horizontalPrintResolution;
			verticalPrintResolution = copy.verticalPrintResolution;
			leftMargin = copy.leftMargin;
			rightMargin = copy.rightMargin;
			topMargin = copy.topMargin;
			bottomMargin = copy.bottomMargin;
			fitToPages = copy.fitToPages;
			password = copy.password;
			passwordHash = copy.passwordHash;
			defaultColumnWidth = copy.defaultColumnWidth;
			defaultRowHeight = copy.defaultRowHeight;
			zoomFactor = copy.zoomFactor;
			showGridLines = copy.showGridLines;
			displayZeroValues = copy.displayZeroValues;
			horizontalFreeze = copy.horizontalFreeze;
			verticalFreeze = copy.verticalFreeze;
			copies = copy.copies;
			header = new HeaderFooter(copy.header);
			footer = new HeaderFooter(copy.footer);
		}
		
		/// <summary> Sets this sheet to be when it is opened in excel</summary>
		public void  setSelected()
		{
			selected = true;
		}
		
		/// <summary> Accessor for the selected nature of the sheet
		/// 
		/// </summary>
		/// <returns> TRUE if this sheet is selected, FALSE otherwise
		/// </returns>
		public bool isSelected()
		{
			return selected;
		}
	}
}
