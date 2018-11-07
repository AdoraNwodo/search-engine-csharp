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
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	/// <summary> A number record.  This is stored as 8 bytes, as opposed to the
	/// 4 byte RK record
	/// </summary>
	public class HyperlinkRecord:RecordData, Hyperlink
	{
		/// <summary> Returns the row number of the top left cell
		/// 
		/// </summary>
		/// <returns> the row number of this cell
		/// </returns>
		virtual public int Row
		{
			get
			{
				return firstRow;
			}
			
		}
		/// <summary> Returns the column number of the top left cell
		/// 
		/// </summary>
		/// <returns> the column number of this cell
		/// </returns>
		virtual public int Column
		{
			get
			{
				return firstColumn;
			}
			
		}
		/// <summary> Returns the row number of the bottom right cell
		/// 
		/// </summary>
		/// <returns> the row number of this cell
		/// </returns>
		virtual public int LastRow
		{
			get
			{
				return lastRow;
			}
			
		}
		/// <summary> Returns the column number of the bottom right cell
		/// 
		/// </summary>
		/// <returns> the column number of this cell
		/// </returns>
		virtual public int LastColumn
		{
			get
			{
				return lastColumn;
			}
			
		}
		/// <summary> Gets the range of cells which activate this hyperlink
		/// The get sheet index methods will all return -1, because the
		/// cells will all be present on the same sheet
		/// 
		/// </summary>
		/// <returns> the range of cells which activate the hyperlink
		/// </returns>
		virtual public Range Range
		{
			get
			{
				return range;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The first row</summary>
		private int firstRow;
		/// <summary> The last row</summary>
		private int lastRow;
		/// <summary> The first column</summary>
		private int firstColumn;
		/// <summary> The last column</summary>
		private int lastColumn;
		
		/// <summary> The URL referred to by this hyperlink</summary>
		private Uri url;
		
		/// <summary> The local file referred to by this hyperlink</summary>
		private System.IO.FileInfo file;
		
		/// <summary> The location in this workbook referred to by this hyperlink</summary>
		private string location;
		
		/// <summary> The range of cells which activate this hyperlink</summary>
		private SheetRangeImpl range;
		
		/// <summary> The type of this hyperlink</summary>
		private LinkType linkType;
		
		/// <summary> The excel type of hyperlink</summary>
		private class LinkType
		{
		}
		
		
		private static readonly LinkType urlLink = new LinkType();
		private static readonly LinkType fileLink = new LinkType();
		private static readonly LinkType workbookLink = new LinkType();
		private static readonly LinkType unknown = new LinkType();
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="s">the sheet
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		internal HyperlinkRecord(Record t, Sheet s, WorkbookSettings ws):base(t)
		{
			
			linkType = unknown;
			
			sbyte[] data = getRecord().Data;
			
			// Build up the range of cells occupied by this hyperlink
			firstRow = IntegerHelper.getInt(data[0], data[1]);
			lastRow = IntegerHelper.getInt(data[2], data[3]);
			firstColumn = IntegerHelper.getInt(data[4], data[5]);
			lastColumn = IntegerHelper.getInt(data[6], data[7]);
			range = new SheetRangeImpl(s, firstColumn, firstRow, lastColumn, lastRow);
			
			// Try and determine the type
			if ((data[28] & 0x3) == 0x03)
			{
				linkType = urlLink;
				bool description = (data[28] & 0x14) != 0;
				
				string urlString = null;
				try
				{
					int startpos = 32;
					if (description)
					{
						int descbytes = IntegerHelper.getInt(data[startpos], data[startpos + 1], data[startpos + 2], data[startpos + 3]);
						startpos += descbytes * 2 + 4;
					}
					
					startpos += 16;
					
					// Get the url, ignoring the 0 char at the end
					int bytes = IntegerHelper.getInt(data[startpos], data[startpos + 1], data[startpos + 2], data[startpos + 3]);
					
					urlString = StringHelper.getUnicodeString(data, bytes / 2 - 1, startpos + 4);
					url = new Uri(urlString);
				}
				catch (Exception e)
				{
					System.Text.StringBuilder sb1 = new System.Text.StringBuilder();
					System.Text.StringBuilder sb2 = new System.Text.StringBuilder();
					NExcel.CellReferenceHelper.getCellReference(firstColumn, firstRow, sb1);
					NExcel.CellReferenceHelper.getCellReference(lastColumn, lastRow, sb2);
					sb1.Insert(0, "Exception when parsing URL ");
					sb1.Append('\"').Append(sb2.ToString()).Append("\".  Using default.");
					logger.warn(sb1, e);
					url = new Uri("http://www.sourceforge.net/projects/nexcel");
				}
			}
			else if ((data[28] & 0x01) != 0)
			{
				linkType = fileLink;
				//      boolean description = (data[28] & 0x14) != 0;
				
				try
				{
					int startpos = 48;
					
					// Get the name of the local file, ignoring the zero character at the
					// end
					int upLevelCount = IntegerHelper.getInt(data[startpos], data[startpos + 1]);
					int chars = IntegerHelper.getInt(data[startpos + 2], data[startpos + 3], data[startpos + 4], data[startpos + 5]);
					string fileName = StringHelper.getString(data, chars - 1, startpos + 6, ws);
					
					System.Text.StringBuilder sb = new System.Text.StringBuilder();
					
					for (int i = 0; i < upLevelCount; i++)
					{
						sb.Append("..\\");
					}
					
					sb.Append(fileName);
					
					file = new System.IO.FileInfo(sb.ToString());
				}
				catch (System.Exception e)
				{
					logger.warn("Exception when parsing file " + e.GetType().FullName + ".");
					file = new System.IO.FileInfo(".");
				}
			}
			else if ((data[28] & 0x08) != 0)
			{
				linkType = workbookLink;
				
				int chars = IntegerHelper.getInt(data[32], data[33], data[34], data[35]);
				location = StringHelper.getUnicodeString(data, chars - 1, 36);
			}
			else
			{
				// give up
				logger.warn("Cannot determine link type");
				return ;
			}
		}
		
		/// <summary> Determines whether this is a hyperlink to a file
		/// 
		/// </summary>
		/// <returns> TRUE if this is a hyperlink to a file, FALSE otherwise
		/// </returns>
		public virtual bool isFile()
		{
			return linkType == fileLink;
		}
		
		/// <summary> Determines whether this is a hyperlink to a web resource
		/// 
		/// </summary>
		/// <returns> TRUE if this is a URL
		/// </returns>
		public virtual bool isURL()
		{
			return linkType == urlLink;
		}
		
		/// <summary> Determines whether this is a hyperlink to a location in this workbook
		/// 
		/// </summary>
		/// <returns> TRUE if this is a link to an internal location
		/// </returns>
		public virtual bool isLocation()
		{
			return linkType == workbookLink;
		}
		
		/// <summary> Gets the URL referenced by this Hyperlink
		/// 
		/// </summary>
		/// <returns> the Uri, or NULL if this hyperlink is not a Uri
		/// </returns>
		public virtual Uri getURL()
		{
			return url;
		}
		
		/// <summary> Returns the local file eferenced by this Hyperlink
		/// 
		/// </summary>
		/// <returns> the file, or NULL if this hyperlink is not a file
		/// </returns>
		public virtual System.IO.FileInfo getFile()
		{
			return file;
		}
		
		/// <summary> Exposes the base class method.  This is used when copying hyperlinks
		/// 
		/// </summary>
		/// <returns> the Record data
		/// </returns>
		public override Record getRecord()
		{
			return base.getRecord();
		}
		
		/// <summary> Gets the location referenced by this hyperlink
		/// 
		/// </summary>
		/// <returns> the location
		/// </returns>
		public virtual string Location
		{
			get
			{
				return location;
			}
		}

		static HyperlinkRecord()
		{
			logger = Logger.getLogger(typeof(HyperlinkRecord));
		}
	}
}
