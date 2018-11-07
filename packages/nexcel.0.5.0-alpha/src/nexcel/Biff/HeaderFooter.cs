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
using Logger = common.Logger;
namespace NExcel.Biff
{
	
	/// <summary> Class which represents an Excel header or footer. Information for this
	/// class came from Microsoft Knowledge Base Article 142136 
	/// (previously Q142136).
	/// 
	/// This class encapsulates three internal structures representing the header
	/// or footer contents which appear on the left, right or central part of the 
	/// page
	/// </summary>
	public abstract class HeaderFooter
	{
		/// <summary> Accessor for the contents which appear on the right hand side of the page
		/// 
		/// </summary>
		/// <returns> the right aligned contents
		/// </returns>
		virtual protected internal Contents RightText
		{
			get
			{
				return right;
			}
			
		}
		/// <summary> Accessor for the contents which in the centre of the page
		/// 
		/// </summary>
		/// <returns> the centrally  aligned contents
		/// </returns>
		virtual protected internal Contents CentreText
		{
			get
			{
				return centre;
			}
			
		}
		/// <summary> Accessor for the contents which appear on the left hand side of the page
		/// 
		/// </summary>
		/// <returns> the left aligned contents
		/// </returns>
		virtual protected internal Contents LeftText
		{
			get
			{
				return left;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		// Codes to format text
		
		/// <summary> Turns bold printing on or off</summary>
		private const string BOLD_TOGGLE = "&B";
		
		/// <summary> Turns underline printing on or off</summary>
		private const string UNDERLINE_TOGGLE = "&U";
		
		/// <summary> Turns italic printing on or off</summary>
		private const string ITALICS_TOGGLE = "&I";
		
		/// <summary> Turns strikethrough printing on or off</summary>
		private const string STRIKETHROUGH_TOGGLE = "&S";
		
		/// <summary> Turns double-underline printing on or off</summary>
		private const string DOUBLE_UNDERLINE_TOGGLE = "&E";
		
		/// <summary> Turns superscript printing on or off</summary>
		private const string SUPERSCRIPT_TOGGLE = "&X";
		
		/// <summary> Turns subscript printing on or off</summary>
		private const string SUBSCRIPT_TOGGLE = "&Y";
		
		/// <summary> Turns outline printing on or off (Macintosh only)</summary>
		private const string OUTLINE_TOGGLE = "&O";
		
		/// <summary> Turns shadow printing on or off (Macintosh only)</summary>
		private const string SHADOW_TOGGLE = "&H";
		
		/// <summary> Left-aligns the characters that follow</summary>
		private const string LEFT_ALIGN = "&L";
		
		/// <summary> Centres the characters that follow</summary>
		private const string CENTRE = "&C";
		
		/// <summary> Right-aligns the characters that follow</summary>
		private const string RIGHT_ALIGN = "&R";
		
		// Codes to insert specific data
		
		/// <summary> Prints the page number</summary>
		private const string PAGENUM = "&P";
		
		/// <summary> Prints the total number of pages in the document</summary>
		private const string TOTAL_PAGENUM = "&N";
		
		/// <summary> Prints the current date</summary>
		private const string DATE = "&D";
		
		/// <summary> Prints the current time</summary>
		private const string TIME = "&T";
		
		/// <summary> Prints the name of the workbook</summary>
		private const string WORKBOOK_NAME = "&F";
		
		/// <summary> Prints the name of the worksheet</summary>
		private const string WORKSHEET_NAME = "&A";
		
		/// <summary> The contents - a simple wrapper around a string buffer</summary>
		public class Contents
		{
			private void  InitBlock(HeaderFooter enclosingInstance)
			{
				this.enclosingInstance = enclosingInstance;
			}
			private HeaderFooter enclosingInstance;
			/// <summary> Sets the font of text subsequently appended to this
			/// object.. Previously appended text is not affected.
			/// <p/>
			/// <strong>Note:</strong> no checking is performed to
			/// determine if fontName is a valid font.
			/// 
			/// </summary>
			/// <param name="fontName">name of the font to use
			/// </param>
			virtual protected internal string FontName
			{
				set
				{
					// Font name must be in quotations
					appendInternal("&\"");
					appendInternal(value);
					appendInternal('\"');
				}
				
			}
			public HeaderFooter Enclosing_Instance
			{
				get
				{
					return enclosingInstance;
				}
				
			}
			/// <summary> The buffer containing the header/footer string</summary>
			private System.Text.StringBuilder contents;
			
			/// <summary> The constructor</summary>
			public Contents(HeaderFooter enclosingInstance)
			{
				InitBlock(enclosingInstance);
				contents = new System.Text.StringBuilder();
			}
			
			/// <summary> Constructor used when reading worksheets.  The string contains all
			/// the formatting (but not alignment characters
			/// 
			/// </summary>
			/// <param name="s">the format string
			/// </param>
			public Contents(HeaderFooter enclosingInstance, string s)
			{
				InitBlock(enclosingInstance);
				contents = new System.Text.StringBuilder(s);
			}
			
			/// <summary> Copy constructor
			/// 
			/// </summary>
			/// <param name="copy">the contents to copy
			/// </param>
			public Contents(HeaderFooter enclosingInstance, Contents copy)
			{
				InitBlock(enclosingInstance);
				contents = new System.Text.StringBuilder(copy.getContents());
			}
			
			/// <summary> Retrieves a <code>String</code>ified
			/// version of this object
			/// 
			/// </summary>
			/// <returns> the header string
			/// </returns>
			protected internal virtual string getContents()
			{
				return contents != null?contents.ToString():"";
			}
			
			/// <summary> Internal method which appends the text to the string buffer
			/// 
			/// </summary>
			/// <param name="">txt
			/// </param>
			private void  appendInternal(string txt)
			{
				if (contents == null)
				{
					contents = new System.Text.StringBuilder();
				}
				
				contents.Append(txt);
			}
			
			/// <summary> Internal method which appends the text to the string buffer
			/// 
			/// </summary>
			/// <param name="">ch
			/// </param>
			private void  appendInternal(char ch)
			{
				if (contents == null)
				{
					contents = new System.Text.StringBuilder();
				}
				
				contents.Append(ch);
			}
			
			/// <summary> Appends the text to the string buffer
			/// 
			/// </summary>
			/// <param name="">txt
			/// </param>
			public virtual void  append(string txt)
			{
				appendInternal(txt);
			}
			
			/// <summary> Turns bold printing on or off. Bold printing
			/// is initially off. Text subsequently appended to
			/// this object will be bolded until this method is
			/// called again.
			/// </summary>
			protected internal virtual void  toggleBold()
			{
				appendInternal(NExcel.Biff.HeaderFooter.BOLD_TOGGLE);
			}
			
			/// <summary> Turns underline printing on or off. Underline printing
			/// is initially off. Text subsequently appended to
			/// this object will be underlined until this method is
			/// called again.
			/// </summary>
			protected internal virtual void  toggleUnderline()
			{
				appendInternal(NExcel.Biff.HeaderFooter.UNDERLINE_TOGGLE);
			}
			
			/// <summary> Turns italics printing on or off. Italics printing
			/// is initially off. Text subsequently appended to
			/// this object will be italicized until this method is
			/// called again.
			/// </summary>
			protected internal virtual void  toggleItalics()
			{
				appendInternal(NExcel.Biff.HeaderFooter.ITALICS_TOGGLE);
			}
			
			/// <summary> Turns strikethrough printing on or off. Strikethrough printing
			/// is initially off. Text subsequently appended to
			/// this object will be striked out until this method is
			/// called again.
			/// </summary>
			protected internal virtual void  toggleStrikethrough()
			{
				appendInternal(NExcel.Biff.HeaderFooter.STRIKETHROUGH_TOGGLE);
			}
			
			/// <summary> Turns double-underline printing on or off. Double-underline printing
			/// is initially off. Text subsequently appended to
			/// this object will be double-underlined until this method is
			/// called again.
			/// </summary>
			protected internal virtual void  toggleDoubleUnderline()
			{
				appendInternal(NExcel.Biff.HeaderFooter.DOUBLE_UNDERLINE_TOGGLE);
			}
			
			/// <summary> Turns superscript printing on or off. Superscript printing
			/// is initially off. Text subsequently appended to
			/// this object will be superscripted until this method is
			/// called again.
			/// </summary>
			protected internal virtual void  toggleSuperScript()
			{
				appendInternal(NExcel.Biff.HeaderFooter.SUPERSCRIPT_TOGGLE);
			}
			
			/// <summary> Turns subscript printing on or off. Subscript printing
			/// is initially off. Text subsequently appended to
			/// this object will be subscripted until this method is
			/// called again.
			/// </summary>
			protected internal virtual void  toggleSubScript()
			{
				appendInternal(NExcel.Biff.HeaderFooter.SUBSCRIPT_TOGGLE);
			}
			
			/// <summary> Turns outline printing on or off (Macintosh only).
			/// Outline printing is initially off. Text subsequently appended
			/// to this object will be outlined until this method is
			/// called again.
			/// </summary>
			protected internal virtual void  toggleOutline()
			{
				appendInternal(NExcel.Biff.HeaderFooter.OUTLINE_TOGGLE);
			}
			
			/// <summary> Turns shadow printing on or off (Macintosh only).
			/// Shadow printing is initially off. Text subsequently appended
			/// to this object will be shadowed until this method is
			/// called again.
			/// </summary>
			protected internal virtual void  toggleShadow()
			{
				appendInternal(NExcel.Biff.HeaderFooter.SHADOW_TOGGLE);
			}
			
			/// <summary> Sets the font size of text subsequently appended to this
			/// object. Previously appended text is not affected.
			/// <p/>
			/// Valid point sizes are between 1 and 99 (inclusive). If
			/// size is outside this range, this method returns false
			/// and does not change font size. If size is within this
			/// range, the font size is changed and true is returned. 
			/// 
			/// </summary>
			/// <param name="size">The size in points. Valid point sizes are
			/// between 1 and 99 (inclusive).
			/// </param>
			/// <returns> true if the font size was changed, false if font
			/// size was not changed because 1 > size > 99. 
			/// </returns>
			protected internal virtual bool setFontSize(int size)
			{
				if (size < 1 || size > 99)
				{
					return false;
				}
				
				// A two digit number should be used -- even if the
				// leading number is just a zero.
				string fontSize;
				if (size < 10)
				{
					// single-digit -- make two digit
					fontSize = "0" + size;
				}
				else
				{
					fontSize = System.Convert.ToString(size);
				}
				
				appendInternal('&');
				appendInternal(fontSize);
				return true;
			}
			
			/// <summary> Appends the page number</summary>
			protected internal virtual void  appendPageNumber()
			{
				appendInternal(NExcel.Biff.HeaderFooter.PAGENUM);
			}
			
			/// <summary> Appends the total number of pages</summary>
			protected internal virtual void  appendTotalPages()
			{
				appendInternal(NExcel.Biff.HeaderFooter.TOTAL_PAGENUM);
			}
			
			/// <summary> Appends the current date</summary>
			protected internal virtual void  appendDate()
			{
				appendInternal(NExcel.Biff.HeaderFooter.DATE);
			}
			
			/// <summary> Appends the current time</summary>
			protected internal virtual void  appendTime()
			{
				appendInternal(NExcel.Biff.HeaderFooter.TIME);
			}
			
			/// <summary> Appends the workbook name</summary>
			protected internal virtual void  appendWorkbookName()
			{
				appendInternal(NExcel.Biff.HeaderFooter.WORKBOOK_NAME);
			}
			
			/// <summary> Appends the worksheet name</summary>
			protected internal virtual void  appendWorkSheetName()
			{
				appendInternal(NExcel.Biff.HeaderFooter.WORKSHEET_NAME);
			}
			
			/// <summary> Clears the contents of this portion</summary>
			protected internal virtual void  clear()
			{
				contents = null;
			}
			
			/// <summary> Queries if the contents are empty
			/// 
			/// </summary>
			/// <returns> TRUE if the contents are empty, FALSE otherwise
			/// </returns>
			protected internal virtual bool empty()
			{
				if (contents == null || contents.Length == 0)
				{
					return true;
				}
				else
				{
					return false;
				}
			}
		}
		
		/// <summary> The left aligned header/footer contents</summary>
		private Contents left;
		
		/// <summary> The right aligned header/footer contents</summary>
		private Contents right;
		
		/// <summary> The centrally aligned header/footer contents</summary>
		private Contents centre;
		
		/// <summary> Default constructor.</summary>
		protected internal HeaderFooter()
		{
			left = createContents();
			right = createContents();
			centre = createContents();
		}
		
		/// <summary> Copy constructor
		/// 
		/// </summary>
		/// <param name="c">the item to copy
		/// </param>
		protected internal HeaderFooter(HeaderFooter hf)
		{
			left = createContents(hf.left);
			right = createContents(hf.right);
			centre = createContents(hf.centre);
		}
		
		/// <summary> Constructor used when reading workbooks to separate the left, right
		/// a central part of the strings into their constituent parts
		/// </summary>
		protected internal HeaderFooter(string s)
		{
			if ((System.Object) s == null)
			{
				left = createContents();
				right = createContents();
				centre = createContents();
				return ;
			}
			
			int pos = 0;
			int leftPos = s.IndexOf(LEFT_ALIGN);
			int rightPos = s.IndexOf(RIGHT_ALIGN);
			int centrePos = s.IndexOf(CENTRE);
			
			// Do the left position string
			if (pos == leftPos)
			{
				if (centrePos != - 1)
				{
					left = createContents(s.Substring(pos + 2, (centrePos) - (pos + 2)));
					pos = centrePos;
				}
				else if (rightPos != - 1)
				{
					left = createContents(s.Substring(pos + 2, (rightPos) - (pos + 2)));
					pos = rightPos;
				}
				else
				{
					left = createContents(s.Substring(pos + 2));
					pos = s.Length;
				}
			}
			
			// Do the centrally positioned part of the string.  This is the default
			// if no alignment string is specified
			if (pos == centrePos || (leftPos == - 1 && rightPos == - 1 && centrePos == - 1))
			{
				if (rightPos != - 1)
				{
					centre = createContents(s.Substring(pos + 2, (rightPos) - (pos + 2)));
					pos = rightPos;
				}
				else
				{
					centre = createContents(s.Substring(pos + 2));
					pos = s.Length;
				}
			}
			
			// Do the right positioned part of the string
			if (pos == rightPos)
			{
				right = createContents(s.Substring(pos + 2));
				pos = s.Length;
			}
			
			if (left == null)
			{
				left = createContents();
			}
			
			if (centre == null)
			{
				centre = createContents();
			}
			
			if (right == null)
			{
				right = createContents();
			}
		}
		
		/// <summary> Retrieves a <code>String</code>ified
		/// version of this object
		/// 
		/// </summary>
		/// <returns> the header string
		/// </returns>
		public override string ToString()
		{
			System.Text.StringBuilder hf = new System.Text.StringBuilder();
			if (!left.empty())
			{
				hf.Append(LEFT_ALIGN);
				hf.Append(left.getContents());
			}
			
			if (!centre.empty())
			{
				hf.Append(CENTRE);
				hf.Append(centre.getContents());
			}
			
			if (!right.empty())
			{
				hf.Append(RIGHT_ALIGN);
				hf.Append(right.getContents());
			}
			
			return hf.ToString();
		}
		
		/// <summary> Clears the contents of the header/footer</summary>
		public virtual void  clear()
		{
			left.clear();
			right.clear();
			centre.clear();
		}
		
		/// <summary> Creates internal class of the appropriate type</summary>
		protected internal abstract Contents createContents();
		
		/// <summary> Creates internal class of the appropriate type</summary>
		protected internal abstract Contents createContents(string s);
		
		/// <summary> Creates internal class of the appropriate type</summary>
		protected internal abstract Contents createContents(Contents c);
		static HeaderFooter()
		{
			logger = Logger.getLogger(typeof(HeaderFooter));
		}
	}
}
