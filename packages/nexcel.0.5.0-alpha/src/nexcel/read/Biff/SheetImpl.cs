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
using NExcel;
using NExcel.Biff;
using NExcel.Biff.Drawing;
using NExcel.Format;

namespace NExcel.Read.Biff
{
	
	/// <summary> Represents a sheet within a workbook.  Provides a handle to the individual
	/// cells, or lines of cells (grouped by Row or Column)
	/// In order to simplify this class due to code bloat, the actual reading
	/// logic has been delegated to the SheetReaderClass.  This class' main
	/// responsibility is now to implement the API methods declared in the
	/// Sheet interface
	/// </summary>
	public class SheetImpl : Sheet
	{
		/// <summary> Returns the number of rows in this sheet
		/// 
		/// </summary>
		/// <returns> the number of rows in this sheet
		/// </returns>
		virtual public int Rows
		{
			get
			{
				// just in case this has been cleared, but something else holds
				// a reference to it
				if (cells == null)
				{
					readSheet();
				}
				
				return numRows;
			}
			
		}
		/// <summary> Returns the number of columns in this sheet
		/// 
		/// </summary>
		/// <returns> the number of columns in this sheet
		/// </returns>
		virtual public int Columns
		{
			get
			{
				// just in case this has been cleared, but something else holds
				// a reference to it
				if (cells == null)
				{
					readSheet();
				}
				
				return numCols;
			}
			
		}
		/// <summary> Gets all the column info records
		/// 
		/// </summary>
		/// <returns> the ColumnInfoRecordArray
		/// </returns>
		virtual public ColumnInfoRecord[] ColumnInfos
		{
			get
			{
				// Just chuck all the column infos we have into an array
				ColumnInfoRecord[] infos = new ColumnInfoRecord[columnInfosArray.Count];
				for (int i = 0; i < columnInfosArray.Count; i++)
				{
					infos[i] = (ColumnInfoRecord) columnInfosArray[i];
				}
				
				return infos;
			}
			
		}
		/// <summary> Gets the hyperlinks on this sheet
		/// 
		/// </summary>
		/// <returns> an array of hyperlinks
		/// </returns>
		virtual public Hyperlink[] Hyperlinks
		{
			get
			{
				Hyperlink[] hl = new Hyperlink[hyperlinks.Count];
				
				for (int i = 0; i < hyperlinks.Count; i++)
				{
					hl[i] = (Hyperlink) hyperlinks[i];
				}
				
				return hl;
			}
			
		}
		/// <summary> Gets the cells which have been merged on this sheet
		/// 
		/// </summary>
		/// <returns> an array of range objects
		/// </returns>
		virtual public Range[] MergedCells
		{
			get
			{
				if (mergedCells == null)
				{
					return new Range[0];
				}
				
				return mergedCells;
			}
			
		}
		/// <summary> Gets the non-default rows.  Used when copying spreadsheets
		/// 
		/// </summary>
		/// <returns> an array of row properties
		/// </returns>
		virtual public RowRecord[] RowProperties
		{
			get
			{
				RowRecord[] rp = new RowRecord[rowProperties.Count];
				for (int i = 0; i < rp.Length; i++)
				{
					rp[i] = (RowRecord) rowProperties[i];
				}
				
				return rp;
			}
			
		}
		/// <summary> Gets the row breaks.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> the explicit row breaks
		/// </returns>
		virtual public int[] RowPageBreaks
		{
			get
			{
				return rowBreaks;
			}
			
		}
		/// <summary> Gets the charts.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> the charts on this page
		/// </returns>
		virtual public Chart[] Charts
		{
			get
			{
				Chart[] ch = new Chart[charts.Count];
				
				for (int i = 0; i < ch.Length; i++)
				{
					ch[i] = (Chart) charts[i];
				}
				return ch;
			}
			
		}
		/// <summary> Gets the drawings.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> the drawings on this page
		/// </returns>
		virtual public Drawing[] Drawings
		{
			get
			{
				System.Object[] dr = drawings.ToArray();
				Drawing[] dr2 = new Drawing[dr.Length];
				Array.Copy(dr, 0, dr2, 0, dr.Length);
				return dr2;
			}
			
		}
		/// <summary> Determines whether the sheet is protected
		/// 
		/// </summary>
		/// <returns> whether or not the sheet is protected
		/// </returns>
		/// <deprecated> in favour of the getSettings() api
		/// </deprecated>
		virtual public bool Protected
		{
			get
			{
				return settings.Protected;
			}
			
		}
		/// <summary> Gets the workspace options for this sheet.  Called during the copy
		/// process
		/// 
		/// </summary>
		/// <returns> the workspace options
		/// </returns>
		virtual public WorkspaceInformationRecord WorkspaceOptions
		{
			get
			{
				return workspaceOptions;
			}
			
		}
		/// <summary> Accessor for the sheet settings
		/// 
		/// </summary>
		/// <returns> the settings for this sheet
		/// </returns>
		virtual public SheetSettings Settings
		{
			get
			{
				return settings;
			}
			
		}
		/// <summary> Accessor for the workbook</summary>
		/// <returns>  the workbook
		/// </returns>
		virtual internal WorkbookParser Workbook
		{
			get
			{
				return workbook;
			}
			
		}
		/// <summary> Used when copying sheets in order to determine the type of this sheet
		/// 
		/// </summary>
		/// <returns> the BOF Record
		/// </returns>
		virtual public BOFRecord SheetBof
		{
			get
			{
				return sheetBof;
			}
			
		}
		/// <summary> Accessor for the environment specific print record, invoked when
		/// copying sheets
		/// 
		/// </summary>
		/// <returns> the environment specific print record
		/// </returns>
		virtual public PLSRecord PLS
		{
			get
			{
				return plsRecord;
			}
			
		}
		/// <summary> The excel file</summary>
		private File excelFile;
		/// <summary> A handle to the shared string table</summary>
		private SSTRecord sharedStrings;
		
		/// <summary> A handle to the sheet BOF record, which indicates the stream type</summary>
		private BOFRecord sheetBof;
		
		/// <summary> A handle to the workbook BOF record, which indicates the stream type</summary>
		private BOFRecord workbookBof;
		
		/// <summary> A handle to the formatting records</summary>
		private FormattingRecords formattingRecords;
		
		/// <summary> The name of this sheet</summary>
		private string name;
		
		/// <summary> The  number of rows</summary>
		private int numRows;
		
		/// <summary> The number of columns</summary>
		private int numCols;
		
		/// <summary> The cells</summary>
		private Cell[][] cells;
		
		/// <summary> The start position in the stream of this sheet</summary>
		private int startPosition;
		
		/// <summary> The list of specified (ie. non default) column widths</summary>
		private ColumnInfoRecord[] columnInfos;
		
		/// <summary> The array of row records</summary>
		private RowRecord[] rowRecords;
		
		/// <summary> The list of non-default row properties</summary>
		private ArrayList rowProperties;
		
		/// <summary> An array of column info records.  They are held this way before
		/// they are transferred to the more convenient array
		/// </summary>
		private ArrayList columnInfosArray;
		
		/// <summary> A list of shared formula groups</summary>
		private ArrayList sharedFormulas;
		
		/// <summary> A list of hyperlinks on this page</summary>
		private ArrayList hyperlinks;
		
		/// <summary> A list of charts on this page</summary>
		private ArrayList charts;
		
		/// <summary> A list of drawings on this page</summary>
		private ArrayList drawings;
		
		/// <summary> A list of merged cells on this page</summary>
		private Range[] mergedCells;
		
		/// <summary> Indicates whether the columnInfos array has been initialized</summary>
		private bool columnInfosInitialized;
		
		/// <summary> Indicates whether the rowRecords array has been initialized</summary>
		private bool rowRecordsInitialized;
		
		/// <summary> Indicates whether or not the dates are based around the 1904 date system</summary>
		private bool nineteenFour;
		
		/// <summary> The workspace options</summary>
		private WorkspaceInformationRecord workspaceOptions;
		
		/// <summary> The hidden flag</summary>
		private bool hidden;
		
		/// <summary> The environment specific print record</summary>
		private PLSRecord plsRecord;
		
		/// <summary> The sheet settings</summary>
		private SheetSettings settings;
		
		/// <summary> The horizontal page breaks contained on this sheet</summary>
		private int[] rowBreaks;
		
		/// <summary> A handle to the workbook which contains this sheet.  Some of the records
		/// need this in order to reference external sheets
		/// </summary>
		private WorkbookParser workbook;
		
		/// <summary> A handle to the workbook settings</summary>
		private WorkbookSettings workbookSettings;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="f">the excel file
		/// </param>
		/// <param name="sst">the shared string table
		/// </param>
		/// <param name="fr">formatting records
		/// </param>
		/// <param name="sb">the bof record which indicates the start of the sheet
		/// </param>
		/// <param name="wb">the bof record which indicates the start of the sheet
		/// </param>
		/// <param name="nf">the 1904 flag
		/// </param>
		/// <param name="wp">the workbook which this sheet belongs to
		/// </param>
		/// <exception cref=""> BiffException
		/// </exception>
		internal SheetImpl(File f, SSTRecord sst, FormattingRecords fr, BOFRecord sb, BOFRecord wb, bool nf, WorkbookParser wp)
		{
			excelFile = f;
			sharedStrings = sst;
			formattingRecords = fr;
			sheetBof = sb;
			workbookBof = wb;
			columnInfosArray = new ArrayList();
			sharedFormulas = new ArrayList();
			hyperlinks = new ArrayList();
			rowProperties = new ArrayList(10);
			columnInfosInitialized = false;
			rowRecordsInitialized = false;
			nineteenFour = nf;
			workbook = wp;
			workbookSettings = workbook.Settings;
			
			// Mark the position in the stream, and then skip on until the end
			startPosition = f.Pos;
			
			if (sheetBof.isChart())
			{
				// Set the start pos to include the bof so the sheet reader can handle it
				startPosition -= (sheetBof.Length + 4);
			}
			
			Record r = null;
			int bofs = 1;
			
			while (bofs >= 1)
			{
				r = f.next();
				
				// use this form for quick performance
				if (r.Code == NExcel.Biff.Type.EOF.Value)
				{
					bofs--;
				}
				
				if (r.Code == NExcel.Biff.Type.BOF.Value)
				{
					bofs++;
				}
			}
		}
		
		/// <summary> Returns the cell specified at this row and at this column
		/// 
		/// </summary>
		/// <param name="row">the row number
		/// </param>
		/// <param name="column">the column number
		/// </param>
		/// <returns> the cell at the specified co-ordinates
		/// </returns>
		public virtual Cell getCell(int column, int row)
		{
			// just in case this has been cleared, but something else holds
			// a reference to it
			if (cells == null)
			{
				readSheet();
			}
			
			Cell c = cells[row][column];
			
			if (c == null)
			{
				c = new EmptyCell(column, row);
				cells[row][column] = c;
			}
			
			return c;
		}
		
		/// <summary> Gets the cell whose contents match the string passed in.
		/// If no match is found, then null is returned.  The search is performed
		/// on a row by row basis, so the lower the row number, the more
		/// efficiently the algorithm will perform
		/// 
		/// </summary>
		/// <param name="contents">the string to match
		/// </param>
		/// <returns> the Cell whose contents match the paramter, null if not found
		/// </returns>
		public virtual Cell findCell(string contents)
		{
			Cell cell = null;
			bool found = false;
			
			for (int i = 0; i < Rows && !found; i++)
			{
				Cell[] row = getRow(i);
				for (int j = 0; j < row.Length && !found; j++)
				{
					if (row[j].Contents.Equals(contents))
					{
						cell = row[j];
						found = true;
					}
				}
			}
			
			return cell;
		}
		
		/// <summary> Gets the cell whose contents match the string passed in.
		/// If no match is found, then null is returned.  The search is performed
		/// on a row by row basis, so the lower the row number, the more
		/// efficiently the algorithm will perform.  This method differs
		/// from the findCell methods in that only cells with labels are
		/// queried - all numerical cells are ignored.  This should therefore
		/// improve performance.
		/// 
		/// </summary>
		/// <param name="contents">the string to match
		/// </param>
		/// <returns> the Cell whose contents match the paramter, null if not found
		/// </returns>
		public virtual LabelCell findLabelCell(string contents)
		{
			LabelCell cell = null;
			bool found = false;
			
			for (int i = 0; i < Rows && !found; i++)
			{
				Cell[] row = getRow(i);
				for (int j = 0; j < row.Length && !found; j++)
				{
					if ((row[j].Type == CellType.LABEL || row[j].Type == CellType.STRING_FORMULA) && row[j].Contents.Equals(contents))
					{
						cell = (LabelCell) row[j];
						found = true;
					}
				}
			}
			
			return cell;
		}
		
		/// <summary> Gets all the cells on the specified row.  The returned array will
		/// be stripped of all trailing empty cells
		/// 
		/// </summary>
		/// <param name="row">the rows whose cells are to be returned
		/// </param>
		/// <returns> the cells on the given row
		/// </returns>
		public virtual Cell[] getRow(int row)
		{
			// just in case this has been cleared, but something else holds
			// a reference to it
			if (cells == null)
			{
				readSheet();
			}
			
			// Find the last non-null cell
			bool found = false;
			int col = numCols - 1;
			while (col >= 0 && !found)
			{
				if (cells[row][col] != null)
				{
					found = true;
				}
				else
				{
					col--;
				}
			}
			
			// Only create entries for non-null cells
			Cell[] c = new Cell[col + 1];
			
			for (int i = 0; i <= col; i++)
			{
				c[i] = getCell(i, row);
			}
			return c;
		}
		
		/// <summary> Gets all the cells on the specified column.  The returned array
		/// will be stripped of all trailing empty cells
		/// 
		/// </summary>
		/// <param name="col">the column whose cells are to be returned
		/// </param>
		/// <returns> the cells on the specified column
		/// </returns>
		public virtual Cell[] getColumn(int col)
		{
			// just in case this has been cleared, but something else holds
			// a reference to it
			if (cells == null)
			{
				readSheet();
			}
			
			// Find the last non-null cell
			bool found = false;
			int row = numRows - 1;
			while (row >= 0 && !found)
			{
				if (cells[row][col] != null)
				{
					found = true;
				}
				else
				{
					row--;
				}
			}
			
			// Only create entries for non-null cells
			Cell[] c = new Cell[row + 1];
			
			for (int i = 0; i <= row; i++)
			{
				c[i] = getCell(col, i);
			}
			return c;
		}
		
		/// <summary> Gets the name of this sheet
		/// 
		/// </summary>
		/// <returns> the name of the sheet
		/// </returns>
		public virtual string Name
		{
		get
		{
		return name;
		}
		}
		
		/// <summary> Sets the name of this sheet
		/// 
		/// </summary>
		/// <param name="s">the sheet name
		/// </param>
		internal void  setName(string s)
		{
			name = s;
		}
		
		/// <summary> Determines whether the sheet is hidden
		/// 
		/// </summary>
		/// <returns> whether or not the sheet is hidden
		/// </returns>
		/// <deprecated> in favour of the getSettings function
		/// </deprecated>
		public virtual bool Hidden
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
		
		/// <summary> Gets the column info record for the specified column.  If no
		/// column is specified, null is returned
		/// 
		/// </summary>
		/// <param name="col">the column
		/// </param>
		/// <returns> the ColumnInfoRecord if specified, NULL otherwise
		/// </returns>
		public virtual ColumnInfoRecord getColumnInfo(int col)
		{
			if (!columnInfosInitialized)
			{
				// Initialize the array
				foreach(ColumnInfoRecord cir in columnInfosArray)
				{
				
				int startcol = Math.Max(0, cir.StartColumn);
				int endcol = Math.Min(columnInfos.Length - 1, cir.EndColumn);
				
				for (int c = startcol; c <= endcol; c++)
				{
				columnInfos[c] = cir;
				}
				
				if (endcol < startcol)
				{
				columnInfos[startcol] = cir;
				}
				}
				
				columnInfosInitialized = true;
			}
			
			return col < columnInfos.Length?columnInfos[col]:null;
		}
		
//		/// <summary> Sets the visibility of this sheet
//		/// 
//		/// </summary>
//		/// <param name="h">hidden flag
//		/// </param>
//		internal void  setHidden(bool h)
//		{
//			hidden = h;
//		}
		
		/// <summary> Clears out the array of cells.  This is done for memory allocation
		/// reasons when reading very large sheets
		/// </summary>
		internal void  clear()
		{
			cells = null;
			mergedCells = null;
			columnInfosArray.Clear();
			sharedFormulas.Clear();
			hyperlinks.Clear();
			columnInfosInitialized = false;
			
			if (!workbookSettings.GCDisabled)
			{
				System.GC.Collect();
			}
		}
		
		/// <summary> Reads in the contents of this sheet</summary>
		internal void  readSheet()
		{
			// If this sheet contains only a chart, then set everything to
			// empty and do not bother parsing the sheet
			// Thanks to steve.brophy for spotting this
			if (!sheetBof.isWorksheet())
			{
				numRows = 0;
				numCols = 0;
				cells = new Cell[0][];
				for (int i = 0; i < 0; i++)
				{
					cells[i] = new Cell[0];
				}
				//      return;
			}
			
			SheetReader reader = new SheetReader(excelFile, sharedStrings, formattingRecords, sheetBof, workbookBof, nineteenFour, workbook, startPosition, this);
			reader.read();
			
			// Take stuff that was read in
			numRows = reader.NumRows;
			numCols = reader.NumCols;
			cells = reader.Cells;
			rowProperties = reader.RowProperties;
			columnInfosArray = reader.ColumnInfosArray;
			hyperlinks = reader.Hyperlinks;
			charts = reader.Charts;
			drawings = reader.Drawings;
			mergedCells = reader.MergedCells;
			settings = reader.Settings;
			settings.Hidden = hidden;
			rowBreaks = reader.RowBreaks;
			workspaceOptions = reader.WorkspaceOptions;
			plsRecord = reader.PLS;
			
			reader = null;
			
			if (!workbookSettings.GCDisabled)
			{
				System.GC.Collect();
			}
			
			if (columnInfosArray.Count > 0)
			{
				ColumnInfoRecord cir = (ColumnInfoRecord) columnInfosArray[columnInfosArray.Count - 1];
				columnInfos = new ColumnInfoRecord[cir.EndColumn + 1];
			}
			else
			{
				columnInfos = new ColumnInfoRecord[0];
			}
		}
		
		/// <summary> Gets the row record.  Usually called by the cell in the specified
		/// row in order to determine its size
		/// 
		/// </summary>
		/// <param name="r">the row
		/// </param>
		/// <returns> the RowRecord for the specified row
		/// </returns>
		internal virtual RowRecord getRowInfo(int r)
		{
			if (!rowRecordsInitialized)
			{
				rowRecords = new RowRecord[Rows];
				
				foreach(RowRecord rr in rowProperties)
				{
				rowRecords[rr.RowNumber] = rr;
				}
			}
			
			return rowRecords[r];
		}
		
		/// <summary> Gets the column format for the specified column
		/// 
		/// </summary>
		/// <param name="col">the column number
		/// </param>
		/// <returns> the column format, or NULL if the column has no specific format
		/// </returns>
		/// <deprecated> use getColumnView instead
		/// </deprecated>
		public virtual NExcel.Format.CellFormat getColumnFormat(int col)
		{
			CellView cv = getColumnView(col);
			return cv.Format;
		}
		
		/// <summary> Gets the column width for the specified column
		/// 
		/// </summary>
		/// <param name="col">the column number
		/// </param>
		/// <returns> the column width, or the default width if the column has no
		/// specified format
		/// </returns>
		public virtual int getColumnWidth(int col)
		{
			return this.getColumnView(col).Size / 256;
		}
		
		/// <summary> Gets the column width for the specified column
		/// 
		/// </summary>
		/// <param name="col">the column number
		/// </param>
		/// <returns> the column format, or the default format if no override is
		/// specified
		/// </returns>
		public virtual CellView getColumnView(int col)
		{
			ColumnInfoRecord cir = getColumnInfo(col);
			CellView cv = new CellView();
			
			if (cir != null)
			{
				cv.Dimension = (cir.Width / 256); //deprecated
				cv.Size = cir.Width;
				cv.Hidden = cir.Hidden;
				cv.Format = formattingRecords.getXFRecord(cir.XFIndex);
			}
			else
			{
				cv.Dimension = (settings.DefaultColumnWidth / 256); //deprecated
				cv.Size = (settings.DefaultColumnWidth);
			}
			
			return cv;
		}
		
		/// <summary> Gets the row height for the specified column
		/// 
		/// </summary>
		/// <param name="row">the row number
		/// </param>
		/// <returns> the row height, or the default height if the row has no
		/// specified format
		/// </returns>
		/// <deprecated> use getRowView instead
		/// </deprecated>
		public virtual int getRowHeight(int row)
		{
			return getRowView(row).Dimension;
		}
		
		/// <summary> Gets the row view for the specified row
		/// 
		/// </summary>
		/// <param name="row">the row number
		/// </param>
		/// <returns> the row format, or the default format if no override is
		/// specified
		/// </returns>
		public virtual CellView getRowView(int row)
		{
			RowRecord rr = getRowInfo(row);
			
			CellView cv = new CellView();
			
			if (rr != null)
			{
				cv.Dimension = (rr.RowHeight); //deprecated
				cv.Size = (rr.RowHeight);
				cv.Hidden = rr.isCollapsed();
			}
			else
			{
				cv.Dimension = (settings.DefaultRowHeight);
				cv.Size = (settings.DefaultRowHeight); //deprecated
			}
			
			return cv;
		}
	}
}
