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
using NExcel.Biff.Drawing;
using NExcel.Format;
namespace NExcel.Read.Biff
{
	
	/// <summary> Reads the sheet.  This functionality was originally part of the
	/// SheetImpl class, but was separated out in order to simplify the former
	/// class
	/// </summary>
	sealed class SheetReader 
	{
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the number of rows
		/// </returns>
		internal int NumRows
		{
			get
			{
				return numRows;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the number of columns
		/// </returns>
		internal int NumCols
		{
			get
			{
				return numCols;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the cells
		/// </returns>
		internal Cell[][] Cells
		{
			get
			{
				return cells;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the row properties
		/// </returns>
		internal ArrayList RowProperties
		{
			get
			{
				return rowProperties;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the column information
		/// </returns>
		internal ArrayList ColumnInfosArray
		{
			get
			{
				return columnInfosArray;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the hyperlinks
		/// </returns>
		internal ArrayList Hyperlinks
		{
			get
			{
				return hyperlinks;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the charts
		/// </returns>
		internal ArrayList Charts
		{
			get
			{
				return charts;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the drawings
		/// </returns>
		internal ArrayList Drawings
		{
			get
			{
				return drawings;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the ranges
		/// </returns>
		internal Range[] MergedCells
		{
			get
			{
				return mergedCells;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the sheet settings
		/// </returns>
		internal SheetSettings Settings
		{
			get
			{
				return settings;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the row breaks
		/// </returns>
		internal int[] RowBreaks
		{
			get
			{
				return rowBreaks;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the workspace options
		/// </returns>
		internal WorkspaceInformationRecord WorkspaceOptions
		{
			get
			{
				return workspaceOptions;
			}
			
		}
		/// <summary> Accessor
		/// 
		/// </summary>
		/// <returns> the environment specific print record
		/// </returns>
		internal PLSRecord PLS
		{
			get
			{
				return plsRecord;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
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
		
		/// <summary> The  number of rows</summary>
		private int numRows;
		
		/// <summary> The number of columns</summary>
		private int numCols;
		
		/// <summary> The cells</summary>
		private Cell[][] cells;
		
		/// <summary> The start position in the stream of this sheet</summary>
		private int startPosition;
		
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
		
		/// <summary> A list of merged cells on this page</summary>
		private Range[] mergedCells;
		
		/// <summary> The list of charts on this page</summary>
		private ArrayList charts;
		
		/// <summary> The list of drawings on this page</summary>
		private ArrayList drawings;
		
		/// <summary> Indicates whether or not the dates are based around the 1904 date system</summary>
		private bool nineteenFour;
		
		/// <summary> The PLS print record</summary>
		private PLSRecord plsRecord;
		
		/// <summary> The workspace options</summary>
		private WorkspaceInformationRecord workspaceOptions;
		
		/// <summary> The horizontal page breaks contained on this sheet</summary>
		private int[] rowBreaks;
		
		/// <summary> The sheet settings</summary>
		private SheetSettings settings;
		
		/// <summary> The workbook settings</summary>
		private WorkbookSettings workbookSettings;
		
		/// <summary> A handle to the workbook which contains this sheet.  Some of the records
		/// need this in order to reference external sheets
		/// </summary>
		private WorkbookParser workbook;
		
		/// <summary> A handle to the sheet</summary>
		private SheetImpl sheet;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="fr">the formatting records
		/// </param>
		/// <param name="sst">the shared string table
		/// </param>
		/// <param name="f">the excel file
		/// </param>
		/// <param name="sb">the bof record which indicates the start of the sheet
		/// </param>
		/// <param name="wb">the bof record which indicates the start of the sheet
		/// </param>
		/// <param name="wp">the workbook which this sheet belongs to
		/// </param>
		/// <param name="sp">the start position of the sheet bof in the excel file
		/// </param>
		/// <param name="sh">the sheet
		/// </param>
		/// <param name="nf">1904 date record flag
		/// </param>
		/// <exception cref=""> BiffException
		/// </exception>
		internal SheetReader(File f, SSTRecord sst, FormattingRecords fr, BOFRecord sb, BOFRecord wb, bool nf, WorkbookParser wp, int sp, SheetImpl sh)
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
			charts = new ArrayList();
			drawings = new ArrayList();
			nineteenFour = nf;
			workbook = wp;
			startPosition = sp;
			sheet = sh;
			settings = new SheetSettings();
			workbookSettings = workbook.Settings;
		}
		
		/// <summary> Adds the cell to the array
		/// 
		/// </summary>
		/// <param name="cell">the cell to add
		/// </param>
		private void  addCell(Cell cell)
		{
			// Sometimes multiple cells (eg. MULBLANK) can exceed the
			// column/row boundaries.  Ignore these
			
			if (cell.Row < numRows && cell.Column < numCols)
			{
				if (cells[cell.Row][cell.Column] != null)
				{
					System.Text.StringBuilder sb = new System.Text.StringBuilder();
					NExcel.CellReferenceHelper.getCellReference(cell.Column, cell.Row, sb);
					logger.warn("Cell " + sb.ToString() + " already contains data");
				}
				cells[cell.Row][cell.Column] = cell;
			}
		}
		
		/// <summary> Reads in the contents of this sheet</summary>
		internal void  read()
		{
			Record r = null;
			BaseSharedFormulaRecord sharedFormula = null;
			bool sharedFormulaAdded = false;
			
			bool cont = true;
			
			// Set the position within the file
			excelFile.Pos = startPosition;
			
			// Handles to the last drawing and obj records
			MsoDrawingRecord msoRecord = null;
			ObjRecord objRecord = null;
			
			// A handle to window2 record
			Window2Record window2Record = null;
			
			// A handle to printgridlines record
			PrintGridLinesRecord printGridLinesRecord = null;
			
			// A handle to printheaders record
			PrintHeadersRecord printHeadersRecord = null;
			
			while (cont)
			{
				r = excelFile.next();
				
				if (r.Type == NExcel.Biff.Type.UNKNOWN && r.Code == 0)
				{
					//System.Console.Error.Write("Warning:  biff code zero found");
					
					// Try a dimension record
					if (r.Length == 0xa)
					{
						logger.warn("Biff code zero found - trying a dimension record.");
						r.Type = (NExcel.Biff.Type.DIMENSION);
					}
					else
					{
						logger.warn("Biff code zero found - Ignoring.");
					}
				}
				
				if (r.Type == NExcel.Biff.Type.DIMENSION)
				{
					DimensionRecord dr = null;
					
					if (workbookBof.isBiff8())
					{
						dr = new DimensionRecord(r);
					}
					else
					{
						dr = new DimensionRecord(r, DimensionRecord.biff7);
					}
					numRows = dr.NumberOfRows;
					numCols = dr.NumberOfColumns;
					cells = new Cell[numRows][];
					for (int i = 0; i < numRows; i++)
					{
						cells[i] = new Cell[numCols];
					}
				}
				else if (r.Type == NExcel.Biff.Type.LABELSST)
				{
					LabelSSTRecord label = new LabelSSTRecord(r, sharedStrings, formattingRecords, sheet);
					addCell(label);
				}
				else if (r.Type == NExcel.Biff.Type.RK || r.Type == NExcel.Biff.Type.RK2)
				{
					RKRecord rkr = new RKRecord(r, formattingRecords, sheet);
					
					if (formattingRecords.isDate(rkr.XFIndex))
					{
						DateCell dc = new DateRecord(rkr, rkr.XFIndex, formattingRecords, nineteenFour, sheet);
						addCell(dc);
					}
					else
					{
						addCell(rkr);
					}
				}
				else if (r.Type == NExcel.Biff.Type.HLINK)
				{
					HyperlinkRecord hr = new HyperlinkRecord(r, sheet, workbookSettings);
					hyperlinks.Add(hr);
				}
				else if (r.Type == NExcel.Biff.Type.MERGEDCELLS)
				{
					MergedCellsRecord mc = new MergedCellsRecord(r, sheet);
					if (mergedCells == null)
					{
						mergedCells = mc.Ranges;
					}
					else
					{
						Range[] newMergedCells = new Range[mergedCells.Length + mc.Ranges.Length];
						Array.Copy(mergedCells, 0, newMergedCells, 0, mergedCells.Length);
						Array.Copy(mc.Ranges, 0, newMergedCells, mergedCells.Length, mc.Ranges.Length);
						mergedCells = newMergedCells;
					}
				}
				else if (r.Type == NExcel.Biff.Type.MULRK)
				{
					MulRKRecord mulrk = new MulRKRecord(r);
					
					// Get the individual cell records from the multiple record
					int num = mulrk.NumberOfColumns;
					int ixf = 0;
					for (int i = 0; i < num; i++)
					{
						ixf = mulrk.getXFIndex(i);
						
						NumberValue nv = new NumberValue(mulrk.Row, mulrk.FirstColumn + i, RKHelper.getDouble(mulrk.getRKNumber(i)), ixf, formattingRecords, sheet);
						
						
						if (formattingRecords.isDate(ixf))
						{
							DateCell dc = new DateRecord(nv, ixf, formattingRecords, nineteenFour, sheet);
							addCell(dc);
						}
						else
						{
							nv.setNumberFormat(formattingRecords.getNumberFormat(ixf));
							addCell(nv);
						}
					}
				}
				else if (r.Type == NExcel.Biff.Type.NUMBER)
				{
					NumberRecord nr = new NumberRecord(r, formattingRecords, sheet);
					
					if (formattingRecords.isDate(nr.XFIndex))
					{
						DateCell dc = new DateRecord(nr, nr.XFIndex, formattingRecords, nineteenFour, sheet);
						addCell(dc);
					}
					else
					{
						addCell(nr);
					}
				}
				else if (r.Type == NExcel.Biff.Type.BOOLERR)
				{
					BooleanRecord br = new BooleanRecord(r, formattingRecords, sheet);
					
					if (br.Error)
					{
						ErrorRecord er = new ErrorRecord(br.getRecord(), formattingRecords, sheet);
						addCell(er);
					}
					else
					{
						addCell(br);
					}
				}
				else if (r.Type == NExcel.Biff.Type.PRINTGRIDLINES)
				{
					printGridLinesRecord = new PrintGridLinesRecord(r);
					settings.PrintGridLines = (printGridLinesRecord.PrintGridLines);
				}
				else if (r.Type == NExcel.Biff.Type.PRINTHEADERS)
				{
					printHeadersRecord = new PrintHeadersRecord(r);
					settings.PrintHeaders = (printHeadersRecord.PrintHeaders);
				}
				else if (r.Type == NExcel.Biff.Type.WINDOW2)
				{
					window2Record = new Window2Record(r);
					
					settings.ShowGridLines = (window2Record.ShowGridLines);
					settings.DisplayZeroValues = (window2Record.DisplayZeroValues);
					settings.setSelected();
				}
				else if (r.Type == NExcel.Biff.Type.PANE)
				{
					PaneRecord pr = new PaneRecord(r);
					
					if (window2Record != null && window2Record.Frozen && window2Record.FrozenNotSplit)
					{
						settings.VerticalFreeze = (pr.RowsVisible);
						settings.HorizontalFreeze = (pr.ColumnsVisible);
					}
				}
				else if (r.Type == NExcel.Biff.Type.CONTINUE)
				{
					;
				}
				else if (r.Type == NExcel.Biff.Type.NOTE)
				{
					;
				}
				else if (r.Type == NExcel.Biff.Type.ARRAY)
				{
					;
				}
				else if (r.Type == NExcel.Biff.Type.PROTECT)
				{
					ProtectRecord pr = new ProtectRecord(r);
					settings.Protected=(pr.IsProtected());
				}
				else if (r.Type == NExcel.Biff.Type.SHAREDFORMULA)
				{
					if (sharedFormula == null)
					{
						logger.warn("Shared template formula is null - " + "trying most recent formula template");
						SharedFormulaRecord lastSharedFormula = (SharedFormulaRecord) sharedFormulas[sharedFormulas.Count - 1];
						
						if (lastSharedFormula != null)
						{
							sharedFormula = lastSharedFormula.TemplateFormula;
						}
					}
					
					SharedFormulaRecord sfr = new SharedFormulaRecord(r, sharedFormula, workbook, workbook, sheet);
					sharedFormulas.Add(sfr);
					sharedFormula = null;
				}
				else if (r.Type == NExcel.Biff.Type.FORMULA || r.Type == NExcel.Biff.Type.FORMULA2)
				{
					FormulaRecord fr = new FormulaRecord(r, excelFile, formattingRecords, workbook, workbook, sheet, workbookSettings);
					
					if (fr.Shared)
					{
						BaseSharedFormulaRecord prevSharedFormula = sharedFormula;
						sharedFormula = (BaseSharedFormulaRecord) fr.Formula;
						
						// See if it fits in any of the shared formulas
						sharedFormulaAdded = addToSharedFormulas(sharedFormula);
						
						if (sharedFormulaAdded)
						{
							sharedFormula = prevSharedFormula;
						}
						
						// If we still haven't added the previous base shared formula,
						// revert it to an ordinary formula and add it to the cell
						if (!sharedFormulaAdded && prevSharedFormula != null)
						{
							// Do nothing.  It's possible for the biff file to contain the
							// record sequence
							// FORMULA-SHRFMLA-FORMULA-SHRFMLA-FORMULA-FORMULA-FORMULA
							// ie. it first lists all the formula templates, then it
							// lists all the individual formulas
							addCell(revertSharedFormula(prevSharedFormula));
						}
					}
					else
					{
						Cell cell = fr.Formula;
						
						// See if the formula evaluates to date
						if (fr.Formula.Type == CellType.NUMBER_FORMULA)
						{
							NumberFormulaRecord nfr = (NumberFormulaRecord) fr.Formula;
							if (formattingRecords.isDate(nfr.XFIndex))
							{
								cell = new DateFormulaRecord(nfr, formattingRecords, workbook, workbook, nineteenFour, sheet);
							}
						}
						
						addCell(cell);
					}
				}
				else if (r.Type == NExcel.Biff.Type.LABEL)
				{
					LabelRecord lr = null;
					
					if (workbookBof.isBiff8())
					{
						lr = new LabelRecord(r, formattingRecords, sheet, workbookSettings);
					}
					else
					{
						lr = new LabelRecord(r, formattingRecords, sheet, workbookSettings, LabelRecord.biff7);
					}
					addCell(lr);
				}
				else if (r.Type == NExcel.Biff.Type.RSTRING)
				{
					RStringRecord lr = null;
					
					// RString records are obsolete in biff 8
					Assert.verify(!workbookBof.isBiff8());
					lr = new RStringRecord(r, formattingRecords, sheet, workbookSettings, RStringRecord.biff7);
					addCell(lr);
				}
				else if (r.Type == NExcel.Biff.Type.NAME)
				{
					;
				}
				else if (r.Type == NExcel.Biff.Type.PASSWORD)
				{
					PasswordRecord pr = new PasswordRecord(r);
					settings.PasswordHash=(pr.PasswordHash);
				}
				else if (r.Type == NExcel.Biff.Type.ROW)
				{
					RowRecord rr = new RowRecord(r);
					
					// See if the row has anything funny about it
					if (!rr.isDefaultHeight() || rr.isCollapsed() || rr.isZeroHeight())
					{
						rowProperties.Add(rr);
					}
				}
				else if (r.Type == NExcel.Biff.Type.BLANK)
				{
					BlankCell bc = new BlankCell(r, formattingRecords, sheet);
					addCell(bc);
				}
				else if (r.Type == NExcel.Biff.Type.MULBLANK)
				{
					MulBlankRecord mulblank = new MulBlankRecord(r);
					
					// Get the individual cell records from the multiple record
					int num = mulblank.NumberOfColumns;
					
					for (int i = 0; i < num; i++)
					{
						int ixf = mulblank.getXFIndex(i);
						
						MulBlankCell mbc = new MulBlankCell(mulblank.Row, mulblank.FirstColumn + i, ixf, formattingRecords, sheet);
						
						addCell(mbc);
					}
				}
				else if (r.Type == NExcel.Biff.Type.SCL)
				{
					SCLRecord scl = new SCLRecord(r);
					settings.ZoomFactor = (scl.ZoomFactor);
				}
				else if (r.Type == NExcel.Biff.Type.COLINFO)
				{
					ColumnInfoRecord cir = new ColumnInfoRecord(r);
					columnInfosArray.Add(cir);
				}
				else if (r.Type == NExcel.Biff.Type.HEADER)
				{
					HeaderRecord hr = null;
					if (workbookBof.isBiff8())
					{
						hr = new HeaderRecord(r, workbookSettings);
					}
					else
					{
						hr = new HeaderRecord(r, workbookSettings, HeaderRecord.biff7);
					}
					
					NExcel.HeaderFooter header = new NExcel.HeaderFooter(hr.Header);
					settings.Header = (header);
				}
				else if (r.Type == NExcel.Biff.Type.FOOTER)
				{
					FooterRecord fr = null;
					if (workbookBof.isBiff8())
					{
						fr = new FooterRecord(r, workbookSettings);
					}
					else
					{
						fr = new FooterRecord(r, workbookSettings, FooterRecord.biff7);
					}
					
					NExcel.HeaderFooter footer = new NExcel.HeaderFooter(fr.Footer);
					settings.Footer=(footer);
				}
				else if (r.Type == NExcel.Biff.Type.SETUP)
				{
					SetupRecord sr = new SetupRecord(r);
					if (sr.isPortrait())
					{
						settings.Orientation=(PageOrientation.PORTRAIT);
					}
					else
					{
						settings.Orientation=(PageOrientation.LANDSCAPE);
					}
					settings.PaperSize = (PaperSize.getPaperSize(sr.PaperSize));
					settings.HeaderMargin = (sr.HeaderMargin);
					settings.FooterMargin = (sr.FooterMargin);
					settings.ScaleFactor = (sr.ScaleFactor);
					settings.PageStart = (sr.PageStart);
					settings.FitWidth = (sr.FitWidth);
					settings.FitHeight = (sr.FitHeight);
					settings.HorizontalPrintResolution = (sr.HorizontalPrintResolution);
					settings.VerticalPrintResolution = (sr.VerticalPrintResolution);
					settings.Copies = (sr.Copies);
					
					if (workspaceOptions != null)
					{
						settings.FitToPages = (workspaceOptions.FitToPages);
					}
				}
				else if (r.Type == NExcel.Biff.Type.WSBOOL)
				{
					workspaceOptions = new WorkspaceInformationRecord(r);
				}
				else if (r.Type == NExcel.Biff.Type.DEFCOLWIDTH)
				{
					DefaultColumnWidthRecord dcwr = new DefaultColumnWidthRecord(r);
					settings.DefaultColumnWidth = (dcwr.Width);
				}
				else if (r.Type == NExcel.Biff.Type.DEFAULTROWHEIGHT)
				{
					DefaultRowHeightRecord drhr = new DefaultRowHeightRecord(r);
					if (drhr.Height != 0)
					{
						settings.DefaultRowHeight= (drhr.Height);
					}
				}
				else if (r.Type == NExcel.Biff.Type.LEFTMARGIN)
				{
					MarginRecord m = new LeftMarginRecord(r);
					settings.LeftMargin=(m.Margin);
				}
				else if (r.Type == NExcel.Biff.Type.RIGHTMARGIN)
				{
					MarginRecord m = new RightMarginRecord(r);
					settings.RightMargin = (m.Margin);
				}
				else if (r.Type == NExcel.Biff.Type.TOPMARGIN)
				{
					MarginRecord m = new TopMarginRecord(r);
					settings.TopMargin=(m.Margin);
				}
				else if (r.Type == NExcel.Biff.Type.BOTTOMMARGIN)
				{
					MarginRecord m = new BottomMarginRecord(r);
					settings.BottomMargin=(m.Margin);
				}
				else if (r.Type == NExcel.Biff.Type.HORIZONTALPAGEBREAKS)
				{
					HorizontalPageBreaksRecord dr = null;
					
					if (workbookBof.isBiff8())
					{
						dr = new HorizontalPageBreaksRecord(r);
					}
					else
					{
						dr = new HorizontalPageBreaksRecord(r, HorizontalPageBreaksRecord.biff7);
					}
					rowBreaks = dr.RowBreaks;
				}
				else if (r.Type == NExcel.Biff.Type.PLS)
				{
					plsRecord = new PLSRecord(r);
				}
				else if (r.Type == NExcel.Biff.Type.OBJ)
				{
					objRecord = new ObjRecord(r);
					
					if (objRecord.Type == ObjRecord.PICTURE && !workbookSettings.DrawingsDisabled)
					{
						if (msoRecord == null)
						{
							logger.warn("object record is not associated with a drawing " + " record - ignoring");
						}
						else
						{
							Drawing drawing = new Drawing(msoRecord, objRecord, workbook.DrawingGroup);
							drawings.Add(drawing);
						}
						msoRecord = null;
						objRecord = null;
					}
				}
				else if (r.Type == NExcel.Biff.Type.MSODRAWING)
				{
					msoRecord = new MsoDrawingRecord(r);
				}
				else if (r.Type == NExcel.Biff.Type.BOF)
				{
					BOFRecord br = new BOFRecord(r);
					Assert.verify(!br.isWorksheet());
					
					int startpos = excelFile.Pos - r.Length - 4;
					
					// Skip to the end of the nested bof
					// Thanks to Rohit for spotting this
					Record r2 = excelFile.next();
					while (r2.Code != NExcel.Biff.Type.EOF.Value)
					{
						r2 = excelFile.next();
					}
					
					if (br.isChart())
					{
						Chart chart = new Chart(msoRecord, objRecord, startpos, excelFile.Pos, excelFile, workbookSettings);
						charts.Add(chart);
						
						if (workbook.DrawingGroup != null)
						{
							workbook.DrawingGroup.add(chart);
						}
						
						// Reset the drawing records
						msoRecord = null;
						objRecord = null;
					}
					
					// If this worksheet is just a chart, then the EOF reached
					// represents the end of the sheet as well as the end of the chart
					if (sheetBof.isChart())
					{
						cont = false;
					}
				}
				else if (r.Type == NExcel.Biff.Type.EOF)
				{
					cont = false;
				}
			}
			
			// Restore the file to its accurate position
			excelFile.restorePos();
			
			// Add all the shared formulas to the sheet as individual formulas
			foreach(SharedFormulaRecord sfr in sharedFormulas)
			{
				Cell[] sfnr = sfr.getFormulas(formattingRecords, nineteenFour);
			
				for (int sf = 0; sf < sfnr.Length; sf++)
				{
					addCell(sfnr[sf]);
				}
			}
			
			// If the last base shared formula wasn't added to the sheet, then
			// revert it to an ordinary formula and add it
			if (!sharedFormulaAdded && sharedFormula != null)
			{
				addCell(revertSharedFormula(sharedFormula));
			}
		}
		
		/// <summary> Sees if the shared formula belongs to any of the shared formula
		/// groups
		/// 
		/// </summary>
		/// <param name="fr">the candidate shared formula
		/// </param>
		/// <returns> TRUE if the formula was added, FALSE otherwise
		/// </returns>
		private bool addToSharedFormulas(BaseSharedFormulaRecord fr)
		{
			bool added = false;
			foreach(SharedFormulaRecord sfr in sharedFormulas)
			{
				if (added) break;
				added = sfr.add(fr);
			}
		
			return added;
		}
		
		/// <summary> Reverts the shared formula passed in to an ordinary formula and adds
		/// it to the list
		/// 
		/// </summary>
		/// <param name="f">the formula
		/// </param>
		/// <returns> the new formula
		/// </returns>
		private Cell revertSharedFormula(BaseSharedFormulaRecord f)
		{
			// String formulas look for a STRING record soon after the formula
			// occurred.  Temporarily the position in the excel file back
			// to the point immediately after the formula record
			int pos = excelFile.Pos;
			excelFile.Pos = f.FilePos;
			
			FormulaRecord fr = new FormulaRecord(f.getRecord(), excelFile, formattingRecords, workbook, workbook, FormulaRecord.ignoreSharedFormula, sheet, workbookSettings);
			
			Cell cell = fr.Formula;
			
			// See if the formula evaluates to date
			if (fr.Formula.Type == CellType.NUMBER_FORMULA)
			{
				NumberFormulaRecord nfr = (NumberFormulaRecord) fr.Formula;
				if (formattingRecords.isDate(fr.XFIndex))
				{
					cell = new DateFormulaRecord(nfr, formattingRecords, workbook, workbook, nineteenFour, sheet);
				}
			}
			
			excelFile.Pos = pos;
			return cell;
		}
		static SheetReader()
		{
			logger = Logger.getLogger(typeof(SheetReader));
		}
	}
}
