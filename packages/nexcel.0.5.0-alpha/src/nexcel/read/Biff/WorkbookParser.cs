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
using common;
using NExcel;
using NExcel.Biff;
using NExcel.Biff.Formula;
using NExcel.Biff.Drawing;
namespace NExcel.Read.Biff
{
	
	/// <summary> Parses the biff file passed in, and builds up an internal representation of
	/// the spreadsheet
	/// </summary>
	public class WorkbookParser:Workbook, ExternalSheet, WorkbookMethods
	{
		/// <summary> Gets the sheets within this workbook.
		/// NOTE:  Use of this method for
		/// very large worksheets can cause performance and out of memory problems.
		/// Use the alternative method getSheet() to retrieve each sheet individually
		/// 
		/// </summary>
		/// <returns> an array of the individual sheets
		/// </returns>
		override public Sheet[] Sheets
		{
			get
			{
				Sheet[] sheetArray = new Sheet[NumberOfSheets];
				
				for (int i = 0; i < NumberOfSheets; i++)
				{
					sheetArray[i] = (Sheet) sheets[i];
				}
				return sheetArray;
			}
			
		}
		/// <summary> Gets the sheet names
		/// 
		/// </summary>
		/// <returns> an array of strings containing the sheet names
		/// </returns>
		override public string[] SheetNames
		{
			get
			{
				string[] names = new string[boundsheets.Count];
				
				BoundsheetRecord br = null;
				for (int i = 0; i < names.Length; i++)
				{
					br = (BoundsheetRecord) boundsheets[i];
					names[i] = br.Name;
				}
				
				return names;
			}
			
		}
		/// <summary> Returns the number of sheets in this workbook
		/// 
		/// </summary>
		/// <returns> the number of sheets in this workbook
		/// </returns>
		override public int NumberOfSheets
		{
			get
			{
				return sheets.Count;
			}
			
		}
		/// <summary> Accessor for the formattingRecords, used by the WritableWorkbook
		/// when creating a copy of this
		/// 
		/// </summary>
		/// <returns> the formatting records
		/// </returns>
		virtual public FormattingRecords FormattingRecords
		{
			get
			{
				return formattingRecords;
			}
			
		}
		/// <summary> Accessor for the externSheet, used by the WritableWorkbook
		/// when creating a copy of this
		/// 
		/// </summary>
		/// <returns> the external sheet record
		/// </returns>
		virtual public ExternalSheetRecord ExternalSheetRecord
		{
			get
			{
				return externSheet;
			}
			
		}
		/// <summary> Accessor for the MsoDrawingGroup, used by the WritableWorkbook
		/// when creating a copy of this
		/// 
		/// </summary>
		/// <returns> the Mso Drawing Group record
		/// </returns>
		virtual public MsoDrawingGroupRecord MsoDrawingGroupRecord
		{
			get
			{
				return msoDrawingGroup;
			}
			
		}
		/// <summary> Accessor for the supbook records, used by the WritableWorkbook
		/// when creating a copy of this
		/// 
		/// </summary>
		/// <returns> the supbook records
		/// </returns>
		virtual public SupbookRecord[] SupbookRecords
		{
			get
			{
				SupbookRecord[] sr = new SupbookRecord[supbooks.Count];
				
				for (int i = 0; i < sr.Length; i++)
				{
					sr[i] = (SupbookRecord) supbooks[i];
				}
				
				return sr;
			}
			
		}
		/// <summary> Accessor for the name records.  Used by the WritableWorkbook when
		/// creating a copy of this
		/// 
		/// </summary>
		/// <returns> the array of names
		/// </returns>
		virtual public NameRecord[] NameRecords
		{
			get
			{
				NameRecord[] na = new NameRecord[nameTable.Count];
				
				for (int i = 0; i < nameTable.Count; i++)
				{
					na[i] = (NameRecord) nameTable[i];
				}
				
				return na;
			}
			
		}
		/// <summary> Accessor for the fonts, used by the WritableWorkbook
		/// when creating a copy of this
		/// </summary>
		/// <returns> the fonts used in this workbook
		/// </returns>
		virtual public Fonts Fonts
		{
			get
			{
				return fonts;
			}
			
		}
		/// <summary> Gets the named ranges
		/// 
		/// </summary>
		/// <returns> the list of named cells within the workbook
		/// </returns>
		override public string[] RangeNames
		{
			get
			{
				ArrayList keylist = new ArrayList(namedRecords.Keys);
				System.Object[] keys = keylist.ToArray();
				string[] names = new string[keys.Length];
				Array.Copy(keys, 0, names, 0, keys.Length);
				
				return names;
			}
			
		}
		/// <summary> Method used when parsing formulas to make sure we are trying
		/// to parse a supported biff version
		/// 
		/// </summary>
		/// <returns> the BOF record
		/// </returns>
		virtual public BOFRecord WorkbookBof
		{
			get
			{
				return workbookBof;
			}
			
		}
		/// <summary> Determines whether the sheet is protected
		/// 
		/// </summary>
		/// <returns> whether or not the sheet is protected
		/// </returns>
		override public bool Protected
		{
			get
			{
				return wbProtected;
			}
			
		}
		/// <summary> Accessor for the settings
		/// 
		/// </summary>
		/// <returns> the workbook settings
		/// </returns>
		virtual public WorkbookSettings Settings
		{
			get
			{
				return settings;
			}
			
		}
		/// <summary> Accessor for the drawing group</summary>
		virtual public DrawingGroup DrawingGroup
		{
			get
			{
				return drawingGroup;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The excel file</summary>
		private File excelFile;
		/// <summary> The number of open bofs</summary>
		private int bofs;
		/// <summary> Indicates whether or not the dates are based around the 1904 date system</summary>
		private bool nineteenFour;
		/// <summary> The shared string table</summary>
		private SSTRecord sharedStrings;
		/// <summary> The names of all the worksheets</summary>
		private ArrayList boundsheets;
		/// <summary> The xf records</summary>
		private FormattingRecords formattingRecords;
		/// <summary> The fonts used by this workbook</summary>
		private Fonts fonts;
		
		/// <summary> The sheets contained in this workbook</summary>
		private ArrayList sheets;
		
		/// <summary> The last sheet accessed</summary>
		private SheetImpl lastSheet;
		
		/// <summary> The index of the last sheet retrieved</summary>
		private int lastSheetIndex;
		
		/// <summary> The named records found in this workbook</summary>
		private Hashtable namedRecords;
		
		/// <summary> The list of named records</summary>
		private ArrayList nameTable;
		
		/// <summary> The external sheet record.  Used by formulas, and names</summary>
		private ExternalSheetRecord externSheet;
		
		/// <summary> The list of supporting workbooks - used by formulas</summary>
		private ArrayList supbooks;
		
		/// <summary> The bof record for this workbook</summary>
		private BOFRecord workbookBof;
		
		/// <summary> The Mso Drawing Group record for this workbook</summary>
		private MsoDrawingGroupRecord msoDrawingGroup;
		
		/// <summary> Workbook protected flag</summary>
		private bool wbProtected;
		
		/// <summary> The workbook settings</summary>
		private WorkbookSettings settings;
		
		/// <summary> The drawings contained in this workbook</summary>
		private DrawingGroup drawingGroup;
		
		/// <summary> Constructs this object from the raw excel data
		/// 
		/// </summary>
		/// <param name="f">the excel 97 biff file
		/// </param>
		/// <param name="s">the workbook settings
		/// </param>
		public WorkbookParser(File f, WorkbookSettings s):base()
		{
			excelFile = f;
			boundsheets = new ArrayList(10);
			fonts = new Fonts();
			formattingRecords = new FormattingRecords(fonts);
			sheets = new ArrayList(10);
			supbooks = new ArrayList(10);
			namedRecords = new Hashtable();
			lastSheetIndex = - 1;
			wbProtected = false;
			settings = s;
		}
		
		/// <summary> Interface method from WorkbookMethods - gets the specified
		/// sheet within this workbook
		/// 
		/// </summary>
		/// <param name="index">the zero based index of the required sheet
		/// </param>
		/// <returns> The sheet specified by the index
		/// </returns>
		public virtual Sheet getReadSheet(int index)
		{
			return getSheet(index);
		}
		
		/// <summary> Gets the specified sheet within this workbook
		/// 
		/// </summary>
		/// <param name="index">the zero based index of the required sheet
		/// </param>
		/// <returns> The sheet specified by the index
		/// </returns>
		public override Sheet getSheet(int index)
		{
			// First see if the last sheet index is the same as this sheet index.
			// If so, then the same sheet is being re-requested, so simply
			// return it instead of rereading it
			if ((lastSheet != null) && lastSheetIndex == index)
			{
				return lastSheet;
			}
			
			// Flush out all of the cached data in the last sheet
			if (lastSheet != null)
			{
				lastSheet.clear();
				
				if (!settings.GCDisabled)
				{
					System.GC.Collect();
				}
			}
			
			lastSheet = (SheetImpl) sheets[index];
			lastSheetIndex = index;
			lastSheet.readSheet();
			
			return lastSheet;
		}
		
		/// <summary> Gets the sheet with the specified name from within this workbook
		/// 
		/// </summary>
		/// <param name="name">the sheet name
		/// </param>
		/// <returns> The sheet with the specified name, or null if it is not found
		/// </returns>
		public override Sheet getSheet(string name)
		{
			// Iterate through the boundsheet records
			int pos = 0;
			bool found = false;
			
			foreach(BoundsheetRecord br in boundsheets)
			{
				if (found) break;
			
				if (br.Name.Equals(name))
				{
					found = true;
				}
				else
				{
					pos++;
				}
			}
			
			return found?getSheet(pos):null;
		}
		
		
		/// <summary> Package protected function which gets the real internal sheet index
		/// based upon  the external sheet reference.  This is used for extern sheet
		/// references  which are specified in formulas
		/// 
		/// </summary>
		/// <param name="index">the external sheet reference
		/// </param>
		/// <returns> the actual sheet index
		/// </returns>
		public virtual int getExternalSheetIndex(int index)
		{
			// For biff7, the whole external reference thing works differently
			// Hopefully for our purposes sheet references will all be local
			if (workbookBof.isBiff7())
			{
				return index;
			}
			
			Assert.verify(externSheet != null);
			
			int firstTab = externSheet.getFirstTabIndex(index);
			
			return firstTab;
		}
		
		/// <summary> Package protected function which gets the real internal sheet index
		/// based upon  the external sheet reference.  This is used for extern sheet
		/// references  which are specified in formulas
		/// 
		/// </summary>
		/// <param name="index">the external sheet reference
		/// </param>
		/// <returns> the actual sheet index
		/// </returns>
		public virtual int getLastExternalSheetIndex(int index)
		{
			// For biff7, the whole external reference thing works differently
			// Hopefully for our purposes sheet references will all be local
			if (workbookBof.isBiff7())
			{
				return index;
			}
			
			Assert.verify(externSheet != null);
			
			int lastTab = externSheet.getLastTabIndex(index);
			
			return lastTab;
		}
		
		/// <summary> Gets the name of the external sheet specified by the index
		/// 
		/// </summary>
		/// <param name="index">the external sheet index
		/// </param>
		/// <returns> the name of the external sheet
		/// </returns>
		public virtual string getExternalSheetName(int index)
		{
			// For biff7, the whole external reference thing works differently
			// Hopefully for our purposes sheet references will all be local
			if (workbookBof.isBiff7())
			{
				BoundsheetRecord br = (BoundsheetRecord) boundsheets[index];
				
				return br.Name;
			}
			
			int supbookIndex = externSheet.getSupbookIndex(index);
			SupbookRecord sr = (SupbookRecord) supbooks[supbookIndex];
			
			int firstTab = externSheet.getFirstTabIndex(index);
			
			if (sr.Type == SupbookRecord.INTERNAL)
			{
				// It's an internal reference - get the name from the boundsheets list
				BoundsheetRecord br = (BoundsheetRecord) boundsheets[firstTab];
				
				return br.Name;
			}
			else if (sr.Type == SupbookRecord.EXTERNAL)
			{
				// External reference - get the sheet name from the supbook record
				System.Text.StringBuilder sb = new System.Text.StringBuilder();
				sb.Append('[');
				sb.Append(sr.FileName);
				sb.Append(']');
				sb.Append(sr.getSheetName(firstTab));
				return sb.ToString();
			}
			
			// An unknown supbook - return unkown
			return "[UNKNOWN]";
		}
		
		/// <summary> Gets the name of the external sheet specified by the index
		/// 
		/// </summary>
		/// <param name="index">the external sheet index
		/// </param>
		/// <returns> the name of the external sheet
		/// </returns>
		public virtual string getLastExternalSheetName(int index)
		{
			// For biff7, the whole external reference thing works differently
			// Hopefully for our purposes sheet references will all be local
			if (workbookBof.isBiff7())
			{
				BoundsheetRecord br = (BoundsheetRecord) boundsheets[index];
				
				return br.Name;
			}
			
			int supbookIndex = externSheet.getSupbookIndex(index);
			SupbookRecord sr = (SupbookRecord) supbooks[supbookIndex];
			
			int lastTab = externSheet.getLastTabIndex(index);
			
			if (sr.Type == SupbookRecord.INTERNAL)
			{
				// It's an internal reference - get the name from the boundsheets list
				BoundsheetRecord br = (BoundsheetRecord) boundsheets[lastTab];
				
				return br.Name;
			}
			else if (sr.Type == SupbookRecord.EXTERNAL)
			{
				// External reference - get the sheet name from the supbook record
				System.Text.StringBuilder sb = new System.Text.StringBuilder();
				sb.Append('[');
				sb.Append(sr.FileName);
				sb.Append(']');
				sb.Append(sr.getSheetName(lastTab));
				return sb.ToString();
			}
			
			// An unknown supbook - return unkown
			return "[UNKNOWN]";
		}
		
		/// <summary> Closes this workbook, and frees makes any memory allocated available
		/// for garbage collection
		/// </summary>
		public override void  close()
		{
			if (lastSheet != null)
			{
				lastSheet.clear();
			}
			excelFile.clear();
			
			if (!settings.GCDisabled)
			{
				System.GC.Collect();
			}
		}
		
		/// <summary> Adds the sheet to the end of the array
		/// 
		/// </summary>
		/// <param name="s">the sheet to add
		/// </param>
		public void  addSheet(Sheet s)
		{
			sheets.Add(s);
		}
		
		/// <summary> Does the hard work of building up the object graph from the excel bytes
		/// 
		/// </summary>
		/// <exception cref=""> BiffException
		/// </exception>
		/// <exception cref=""> PasswordException if the workbook is password protected
		/// </exception>
		protected internal override void  parse()
		{
			Record r = null;
			
			BOFRecord bof = new BOFRecord(excelFile.next());
			workbookBof = bof;
			bofs++;
			
			if (!bof.isBiff8() && !bof.isBiff7())
			{
				throw new BiffException(BiffException.unrecognizedBiffVersion);
			}
			
			if (!bof.isWorkbookGlobals())
			{
				throw new BiffException(BiffException.expectedGlobals);
			}
			ArrayList continueRecords = new ArrayList();
			nameTable = new ArrayList();
			
			// Skip to the first worksheet
			while (bofs == 1)
			{
				r = excelFile.next();
				
				if (r.Type == NExcel.Biff.Type.SST)
				{
					continueRecords.Clear();
					Record nextrec = excelFile.peek();
					while (nextrec.Type == NExcel.Biff.Type.CONTINUE)
					{
						continueRecords.Add(excelFile.next());
						nextrec = excelFile.peek();
					}
					
					// cast the array
					System.Object[] rec = continueRecords.ToArray();
					Record[] records = new Record[rec.Length];
					Array.Copy(rec, 0, records, 0, rec.Length);
					
					sharedStrings = new SSTRecord(r, records, settings);
				}
				else if (r.Type == NExcel.Biff.Type.FILEPASS)
				{
					throw new PasswordException();
				}
				else if (r.Type == NExcel.Biff.Type.NAME)
				{
					NameRecord nr = null;
					
					if (bof.isBiff8())
					{
						nr = new NameRecord(r, settings, namedRecords.Count);
					}
					else
					{
						nr = new NameRecord(r, settings, namedRecords.Count, NameRecord.biff7);
					}
					
					namedRecords[nr.Name] =  nr;
					nameTable.Add(nr);
				}
				else if (r.Type == NExcel.Biff.Type.FONT)
				{
					FontRecord fr = null;
					
					if (bof.isBiff8())
					{
						fr = new FontRecord(r, settings);
					}
					else
					{
						fr = new FontRecord(r, settings, FontRecord.biff7);
					}
					fonts.addFont(fr);
				}
				else if (r.Type == NExcel.Biff.Type.PALETTE)
				{
					NExcel.Biff.PaletteRecord palette = new NExcel.Biff.PaletteRecord(r);
					formattingRecords.Palette = palette;
				}
				else if (r.Type == NExcel.Biff.Type.NINETEENFOUR)
				{
					NineteenFourRecord nr = new NineteenFourRecord(r);
					nineteenFour = nr.is1904();
				}
				else if (r.Type == NExcel.Biff.Type.FORMAT)
				{
					FormatRecord fr = null;
					if (bof.isBiff8())
					{
						fr = new FormatRecord(r, settings, FormatRecord.biff8);
					}
					else
					{
						fr = new FormatRecord(r, settings, FormatRecord.biff7);
					}
					try
					{
						formattingRecords.addFormat(fr);
					}
					catch (NumFormatRecordsException e)
					{
						// This should not happen.  Bomb out
						//          Assert.verify(false, e.getMessage());
						Assert.verify(false, "This should not happen. 64");
					}
				}
				else if (r.Type == NExcel.Biff.Type.XF)
				{
					XFRecord xfr = null;
					if (bof.isBiff8())
					{
						xfr = new XFRecord(r, XFRecord.biff8);
					}
					else
					{
						xfr = new XFRecord(r, XFRecord.biff7);
					}
					
					try
					{
						formattingRecords.addStyle(xfr);
					}
					catch (NumFormatRecordsException e)
					{
						// This should not happen.  Bomb out
						//          Assert.verify(false, e.getMessage());
						Assert.verify(false, "This should not happen. 59");
					}
				}
				else if (r.Type == NExcel.Biff.Type.BOUNDSHEET)
				{
					BoundsheetRecord br = null;
					
					if (bof.isBiff8())
					{
						br = new BoundsheetRecord(r);
					}
					else
					{
						br = new BoundsheetRecord(r, BoundsheetRecord.biff7);
					}
					
					if (br.isSheet() || br.Chart)
					{
						boundsheets.Add(br);
					}
				}
				else if (r.Type == NExcel.Biff.Type.EXTERNSHEET)
				{
					if (bof.isBiff8())
					{
						externSheet = new ExternalSheetRecord(r, settings);
					}
					else
					{
						externSheet = new ExternalSheetRecord(r, settings, ExternalSheetRecord.biff7);
					}
				}
				else if (r.Type == NExcel.Biff.Type.CODEPAGE)
				{
					CodepageRecord cr = new CodepageRecord(r);
					settings.CharacterSet = cr.CharacterSet;
				}
				else if (r.Type == NExcel.Biff.Type.SUPBOOK)
				{
					SupbookRecord sr = new SupbookRecord(r, settings);
					supbooks.Add(sr);
				}
				else if (r.Type == NExcel.Biff.Type.PROTECT)
				{
					ProtectRecord pr = new ProtectRecord(r);
					wbProtected = pr.IsProtected();
				}
				else if (r.Type == NExcel.Biff.Type.MSODRAWINGGROUP)
				{
					msoDrawingGroup = new MsoDrawingGroupRecord(r);
					
					if (drawingGroup == null)
					{
						drawingGroup = new DrawingGroup(DrawingGroup.READ);
					}
					
					drawingGroup.add(msoDrawingGroup);
					
					Record nextrec = excelFile.peek();
					while (nextrec.Type == NExcel.Biff.Type.CONTINUE)
					{
						drawingGroup.add(excelFile.next());
						nextrec = excelFile.peek();
					}
				}
				else if (r.Type == NExcel.Biff.Type.EOF)
				{
					bofs--;
				}
			}
			
			bof = null;
			if (excelFile.hasNext())
			{
				r = excelFile.next();
				
				if (r.Type == NExcel.Biff.Type.BOF)
				{
					bof = new BOFRecord(r);
				}
			}
			
			// Only get sheets for which there is a corresponding Boundsheet record
			while (bof != null && NumberOfSheets < boundsheets.Count)
			{
				if (!bof.isBiff8() && !bof.isBiff7())
				{
					throw new BiffException(BiffException.unrecognizedBiffVersion);
				}
				
				if (bof.isWorksheet())
				{
					// Read the sheet in
					SheetImpl s = new SheetImpl(excelFile, sharedStrings, formattingRecords, bof, workbookBof, nineteenFour, this);
					
					BoundsheetRecord br = (BoundsheetRecord) boundsheets[NumberOfSheets];
					s.setName(br.Name);
					s.Hidden = br.isHidden();
					addSheet(s);
				}
				else if (bof.isChart())
				{
					// Read the sheet in
					SheetImpl s = new SheetImpl(excelFile, sharedStrings, formattingRecords, bof, workbookBof, nineteenFour, this);
					
					BoundsheetRecord br = (BoundsheetRecord) boundsheets[NumberOfSheets];
					s.setName(br.Name);
					s.Hidden = br.isHidden();
					addSheet(s);
				}
				else
				{
					logger.warn("BOF is unrecognized");
					
					
					while (excelFile.hasNext() && r.Type != NExcel.Biff.Type.EOF)
					{
						r = excelFile.next();
					}
				}
				
				// The next record will normally be a BOF or empty padding until
				// the end of the block is reached.  In exceptionally unlucky cases,
				// the last EOF  will coincide with a block division, so we have to
				// check there is more data to retrieve.
				// Thanks to liamg for spotting this
				bof = null;
				if (excelFile.hasNext())
				{
					r = excelFile.next();
					
					if (r.Type == NExcel.Biff.Type.BOF)
					{
						bof = new BOFRecord(r);
					}
				}
			}
		}
		
		/// <summary> Gets the named cell from this workbook.  If the name refers to a
		/// range of cells, then the cell on the top left is returned.  If
		/// the name cannot be found, null is returned
		/// 
		/// </summary>
		/// <param name="name">the name of the cell/range to search for
		/// </param>
		/// <returns> the cell in the top left of the range if found, NULL
		/// otherwise
		/// </returns>
		public override Cell findCellByName(string name)
		{
			NameRecord nr = (NameRecord) namedRecords[name];
			
			if (nr == null)
			{
				return null;
			}
			
			NameRecord.NameRange[] ranges = nr.Ranges;
			
			// Go and retrieve the first cell in the first range
			Sheet s = getSheet(ranges[0].ExternalSheet);
			Cell cell = s.getCell(ranges[0].FirstColumn, ranges[0].FirstRow);
			
			return cell;
		}
		
		/// <summary> Gets the named range from this workbook.  The Range object returns
		/// contains all the cells from the top left to the bottom right
		/// of the range.
		/// If the named range comprises an adjacent range,
		/// the Range[] will contain one object; for non-adjacent
		/// ranges, it is necessary to return an array of .Length greater than
		/// one.
		/// If the named range contains a single cell, the top left and
		/// bottom right cell will be the same cell
		/// 
		/// </summary>
		/// <param name="name">the name to find
		/// </param>
		/// <returns> the range of cells
		/// </returns>
		public override Range[] findByName(string name)
		{
			NameRecord nr = (NameRecord) namedRecords[name];
			
			if (nr == null)
			{
				return null;
			}
			
			NameRecord.NameRange[] ranges = nr.Ranges;
			
			Range[] cellRanges = new Range[ranges.Length];
			
			for (int i = 0; i < ranges.Length; i++)
			{
				cellRanges[i] = new RangeImpl(this, getExternalSheetIndex(ranges[i].ExternalSheet), ranges[i].FirstColumn, ranges[i].FirstRow, getLastExternalSheetIndex(ranges[i].ExternalSheet), ranges[i].LastColumn, ranges[i].LastRow);
			}
			
			return cellRanges;
		}
		
		/// <summary> Accessor/implementation method for the external sheet reference
		/// 
		/// </summary>
		/// <param name="sheetName">the sheet name to look for
		/// </param>
		/// <returns> the external sheet index
		/// </returns>
		public virtual int getExternalSheetIndex(string sheetName)
		{
			return 0;
		}
		
		/// <summary> Accessor/implementation method for the external sheet reference
		/// 
		/// </summary>
		/// <param name="sheetName">the sheet name to look for
		/// </param>
		/// <returns> the external sheet index
		/// </returns>
		public virtual int getLastExternalSheetIndex(string sheetName)
		{
			return 0;
		}
		
		/// <summary> Gets the name at the specified index
		/// 
		/// </summary>
		/// <param name="index">the index into the name table
		/// </param>
		/// <returns> the name of the cell
		/// </returns>
		public virtual string getName(int index)
		{
			Assert.verify(index >= 0 && index < nameTable.Count);
			return ((NameRecord) nameTable[index]).Name;
		}
		
		/// <summary> Gets the index of the name record for the name
		/// 
		/// </summary>
		/// <param name="name">the name to search for
		/// </param>
		/// <returns> the index in the name table
		/// </returns>
		public virtual int getNameIndex(string name)
		{
			NameRecord nr = (NameRecord) namedRecords[name];
			
			return nr != null?nr.Index:0;
		}
		static WorkbookParser()
		{
			logger = Logger.getLogger(typeof(WorkbookParser));
		}
	}
}
