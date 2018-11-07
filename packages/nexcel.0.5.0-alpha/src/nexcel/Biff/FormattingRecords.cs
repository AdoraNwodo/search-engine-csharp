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
using System.Collections;
using NExcelUtils;
using common;
// [TODO-NExcel_Next]
//import NExcel.Write.biff.File;
using NExcel.Format;
namespace NExcel.Biff
{
	
	/// <summary> The list of XF records and formatting records for the workbook</summary>
	public class FormattingRecords
	{
		/// <summary> Accessor for the fonts used by this workbook
		/// 
		/// </summary>
		/// <returns> the fonts container
		/// </returns>
		virtual protected internal Fonts Fonts
		{
			// [TODO-NExcel_Next]
			//  /**
			//   * Writes out all the format records and the XF records
			//   *
			//   * @param outputFile the file to write to
			//   * @exception IOException
			//   */
			//  public void write(File outputFile) throws IOException
			//  {
			//    // Write out all the formats
			//    Iterator i = formatsList.iterator();
			//    FormatRecord fr = null;
			//    while (i.hasNext())
			//    {
			//      fr = (FormatRecord) i.next();
			//      outputFile.write(fr);
			//    }
			//
			//    // Write out the styles
			//    i = xfRecords.iterator();
			//    XFRecord xfr = null;
			//    while (i.hasNext())
			//    {
			//      xfr = (XFRecord) i.next();
			//      outputFile.write(xfr);
			//    }
			//
			//    // Write out the style records
			//    BuiltInStyle style = new BuiltInStyle(0x10, 3);
			//    outputFile.write(style);
			//
			//    style = new BuiltInStyle(0x11, 6);
			//    outputFile.write(style);
			//
			//    style = new BuiltInStyle(0x12, 4);
			//    outputFile.write(style);
			//
			//    style = new BuiltInStyle(0x13, 7);
			//    outputFile.write(style);
			//
			//    style = new BuiltInStyle(0x0, 0);
			//    outputFile.write(style);
			//
			//    style = new BuiltInStyle(0x14, 5);
			//    outputFile.write(style);
			//  }
			//
			
			get
			{
				return fonts;
			}
			
		}
		/// <summary> Gets the number of formatting records on the list.  This is used by the
		/// writable subclass because there is an upper limit on the amount of
		/// format records that are allowed to be present in an excel sheet
		/// 
		/// </summary>
		/// <returns> the number of format records present
		/// </returns>
		virtual protected internal int NumberOfFormatRecords
		{
			get
			{
				return formatsList.Count;
			}
			
		}
		/// <summary> Accessor for the colour palette
		/// 
		/// </summary>
		/// <returns> the palette
		/// </returns>
		/// <summary> Called from the WorkbookParser to set the colour palette
		/// 
		/// </summary>
		/// <param name="pr">the palette
		/// </param>
		virtual public PaletteRecord Palette
		{
			get
			{
				return palette;
			}
			
			set
			{
				palette = value;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> A hash map of FormatRecords, for random access retrieval when reading
		/// in a spreadsheet
		/// </summary>
		private Hashtable formats;
		
		/// <summary> A list of formats, used when writing out a spreadsheet</summary>
		private ArrayList formatsList;
		
		/// <summary> The list of extended format records</summary>
		private ArrayList xfRecords;
		
		/// <summary> The next available index number for custom format records</summary>
		private int nextCustomIndexNumber;
		
		/// <summary> A handle to the available fonts</summary>
		private Fonts fonts;
		
		/// <summary> The colour palette</summary>
		private PaletteRecord palette;
		
		/// <summary> The start index number for custom format records</summary>
		private const int customFormatStartIndex = 0xa4;
		
		/// <summary> The maximum number of format records.  This is some weird internal
		/// Excel constraint
		/// </summary>
		private const int maxFormatRecordsIndex = 0x1b9;
		
		/// <summary> The minimum number of XF records for a sheet.  The rationalization
		/// processes commences immediately after this number
		/// </summary>
		private const int minXFRecords = 21;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="f">the container for the fonts
		/// </param>
		public FormattingRecords(Fonts f)
		{
			xfRecords = new ArrayList(10);
			formats = new Hashtable(10);
			formatsList = new ArrayList(10);
			fonts = f;
			nextCustomIndexNumber = customFormatStartIndex;
		}
		
		/// <summary> Adds an extended formatting record to the list.  If the XF record passed
		/// in has not been initialized, its index is determined based on the
		/// xfRecords list, and
		/// this position is passed to the XF records initialize method
		/// 
		/// </summary>
		/// <param name="xf">the xf record to add
		/// </param>
		/// <exception cref=""> NumFormatRecordsException
		/// </exception>
		public void  addStyle(XFRecord xf)
		{
			if (!xf.isInitialized())
			{
				int pos = xfRecords.Count;
				xf.initialize(pos, this, fonts);
				xfRecords.Add(xf);
			}
			else
			{
				// The XF record has probably been read in.  If the index is greater
				// Than the size of the list, then it is not a preset format,
				// so add it
				if (xf.getXFIndex() >= xfRecords.Count)
				{
					xfRecords.Add(xf);
				}
			}
		}
		
		/// <summary> Adds a cell format to the hash map, keyed on its index.  If the format
		/// record is not initialized, then its index number is determined and its
		/// initialize method called.  If the font is not a built in format, then it
		/// is added to the list of formats for writing out
		/// 
		/// </summary>
		/// <param name="fr">the format record
		/// </param>
		public void  addFormat(DisplayFormat fr)
		{
			if (!fr.isInitialized())
			{
				fr.initialize(nextCustomIndexNumber);
				nextCustomIndexNumber++;
			}
			
			if (nextCustomIndexNumber > maxFormatRecordsIndex)
			{
				nextCustomIndexNumber = maxFormatRecordsIndex;
				throw new NumFormatRecordsException();
			}
			
			if (fr.FormatIndex >= nextCustomIndexNumber)
			{
				nextCustomIndexNumber = fr.FormatIndex + 1;
			}
			
			if (!fr.isBuiltIn())
			{
				formatsList.Add(fr);
				formats[fr.FormatIndex] =  fr;
			}
		}
		
		/// <summary> Sees if the extended formatting record at the specified position
		/// represents a date.  First checks against the built in formats, and
		/// then checks against the hash map of FormatRecords
		/// 
		/// </summary>
		/// <param name="pos">the xf format index
		/// </param>
		/// <returns> TRUE if this format index is formatted as a Date
		/// </returns>
		public bool isDate(int pos)
		{
			XFRecord xfr = (XFRecord) xfRecords[pos];
			
			if (xfr.isDate())
			{
				return true;
			}
			
			FormatRecord fr = (FormatRecord) formats[xfr.FormatRecord];
			
			return fr == null?false:fr.Date;
		}
		
		/// <summary> Gets the DateFormat used to format the cell.
		/// 
		/// </summary>
		/// <param name="pos">the xf format index
		/// </param>
		/// <returns> the DateFormat object used to format the date in the original
		/// excel cell
		/// </returns>
		public DateTimeFormatInfo getDateFormat(int pos)
		{
			XFRecord xfr = (XFRecord) xfRecords[pos];
			
			if (xfr.isDate())
			{
				return xfr.DateFormat;
			}
			
			FormatRecord fr = (FormatRecord) formats[xfr.FormatRecord];
			
			if (fr == null)
			{
				return null;
			}
			
			return fr.Date?fr.DateFormat:null;
		}
		
		/// <summary> Gets the NumberFormatInfo used to format the cell.
		/// 
		/// </summary>
		/// <param name="pos">the xf format index
		/// </param>
		/// <returns> the DateFormat object used to format the date in the original
		/// excel cell
		/// </returns>
		public NumberFormatInfo getNumberFormat(int pos)
		{
			XFRecord xfr = (XFRecord) xfRecords[pos];
			
			if (xfr.isNumber())
			{
				return xfr.NumberFormat;
			}
			
			FormatRecord fr = (FormatRecord) formats[xfr.FormatRecord];
			
			if (fr == null)
			{
				return null;
			}
			
			return fr.Number?fr.NumberFormat:null;
		}
		
		/// <summary> Gets the format record
		/// 
		/// </summary>
		/// <param name="index">the formatting record index to retrieve
		/// </param>
		/// <returns> the format record at the specified index
		/// </returns>
		internal virtual FormatRecord getFormatRecord(int index)
		{
			return (FormatRecord) formats[index];
		}
		
		/// <summary> Gets the XFRecord for the specified index.  Used when copying individual
		/// cells
		/// 
		/// </summary>
		/// <param name="index">the XF record to retrieve
		/// </param>
		/// <returns> the XF record at the specified index
		/// </returns>
		public XFRecord getXFRecord(int index)
		{
			return (XFRecord) xfRecords[index];
		}
		
		/// <summary> Rationalizes all the fonts, removing duplicate entries
		/// 
		/// </summary>
		/// <returns> the list of new font index number
		/// </returns>
		public virtual IndexMapping rationalizeFonts()
		{
			return fonts.rationalize();
		}
		
		/// <summary> Rationalizes the cell formats.  Duplicate
		/// formats are removed and the format indexed of the cells
		/// adjusted accordingly
		/// 
		/// </summary>
		/// <param name="fontMapping">the font mapping index numbers
		/// </param>
		/// <param name="formatMapping">the format mapping index numbers
		/// </param>
		/// <returns> the list of new font index number
		/// </returns>
		public virtual IndexMapping rationalize(IndexMapping fontMapping, IndexMapping formatMapping)
		{
			// Update the index codes for the XF records using the format
			// mapping and the font mapping
			// at the same time
			foreach (XFRecord xfr in xfRecords)
			{
			if (xfr.FormatRecord >= customFormatStartIndex)
			{
			xfr.FormatIndex = formatMapping.getNewIndex(xfr.FormatRecord);
			}
			
			xfr.FontIndex = fontMapping.getNewIndex(xfr.FontIndex);
			}
			
			ArrayList newrecords = new ArrayList(minXFRecords);
			IndexMapping mapping = new IndexMapping(xfRecords.Count);
			int numremoved = 0;
			
			// Copy across the fundamental styles
			for (int i = 0; i < minXFRecords; i++)
			{
				newrecords.Add(xfRecords[i]);
				mapping.setMapping(i, i);
			}
			
			// Iterate through the old list
			for (int i = minXFRecords; i < xfRecords.Count; i++)
			{
				XFRecord xf = (XFRecord) xfRecords[i];
				
				// Compare against formats already on the list
				bool duplicate = false;
				foreach (XFRecord xf2 in newrecords)
				{
				if (duplicate) break;
				
				if (xf2.Equals(xf))
				{
				duplicate = true;
				mapping.setMapping(i, mapping.getNewIndex(xf2.getXFIndex()));
				numremoved++;
				}
				}
				
				// If this format is not a duplicate then add it to the new list
				if (!duplicate)
				{
					newrecords.Add(xf);
					mapping.setMapping(i, i - numremoved);
				}
			}
			
			// It is sufficient to merely change the xf index field on all XFRecords
			// In this case, CellValues which refer to defunct format records
			// will nevertheless be written out with the correct index number
			foreach (XFRecord xf in xfRecords)
			{
			xf.rationalize(mapping);
			}
			
			// Set the new list
			xfRecords = newrecords;
			
			return mapping;
		}
		
		/// <summary> Rationalizes the display formats.  Duplicate
		/// formats are removed and the format indices of the cells
		/// adjusted accordingly.  It is invoked immediately prior to writing
		/// writing out the sheet
		/// </summary>
		/// <returns> the index mapping between the old display formats and the
		/// rationalized ones
		/// </returns>
		public virtual IndexMapping rationalizeDisplayFormats()
		{
			ArrayList newformats = new ArrayList();
			int numremoved = 0;
			IndexMapping mapping = new IndexMapping(nextCustomIndexNumber);
			
			// Iterate through the old list
			//    Iterator i = formatsList.iterator();
//			DisplayFormat df = null;
//			DisplayFormat df2 = null;
			bool duplicate = false;
			foreach (DisplayFormat df in formatsList)
			{
				Assert.verify(!df.isBuiltIn());
			
				// Compare against formats already on the list
				duplicate = false;
				foreach (DisplayFormat df2 in newformats)
				{
					if (duplicate) break;
			
					if (df2.Equals(df))
					{
						duplicate = true;
						mapping.setMapping(df.FormatIndex,
							mapping.getNewIndex(df2.FormatIndex));
						numremoved++;
					}
				}
			
				// If this format is not a duplicate then add it to the new list
				if (!duplicate)
				{
					newformats.Add(df);
					int indexnum = df.FormatIndex - numremoved;
					if (indexnum > maxFormatRecordsIndex)
					{
						logger.warn("Too many number formats - using default format.");
						indexnum = 0; // the default number format index
					}
					mapping.setMapping(df.FormatIndex,
						df.FormatIndex - numremoved);
				}
			}
			
			// Set the new list
			formatsList = newformats;
			
			// Update the index codes for the remaining formats
			foreach(DisplayFormat df in formatsList)
			{
				df.initialize(mapping.getNewIndex(df.FormatIndex));
			}  
			
			return mapping;
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
			if (palette == null)
			{
				palette = new PaletteRecord();
			}
			palette.setColourRGB(c, r, g, b);
		}
		static FormattingRecords()
		{
			logger = Logger.getLogger(typeof(FormattingRecords));
		}
	}
}
