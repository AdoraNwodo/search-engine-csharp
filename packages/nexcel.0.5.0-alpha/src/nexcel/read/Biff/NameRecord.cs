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
namespace NExcel.Read.Biff
{
	
	/// <summary> Holds an excel name record, and the details of the cells/ranges it refers
	/// to
	/// </summary>
	public class NameRecord:RecordData
	{
		/// <summary> Gets the name
		/// 
		/// </summary>
		/// <returns> the strings
		/// </returns>
		virtual public string Name
		{
			get
			{
				return name;
			}
			
		}
		/// <summary> Gets the array of ranges for this name.  This method is public as it is
		/// used from the writable side when copying ranges
		/// 
		/// </summary>
		/// <returns> the ranges
		/// </returns>
		virtual public NameRange[] Ranges
		{
			get
			{
				System.Object[] o = ranges.ToArray();
				NameRange[] nr = new NameRange[o.Length];
				
				for (int i = 0; i < o.Length; i++)
				{
					nr[i] = (NameRange) o[i];
				}
				
				return nr;
			}
			
		}
		/// <summary> Accessor for the index into the name table
		/// 
		/// </summary>
		/// <returns> the 0-based index into the name table
		/// </returns>
		virtual internal int Index
		{
			get
			{
				return index;
			}
			
		}
		/// <summary> Called when copying a sheet.  Just returns the raw data
		/// 
		/// </summary>
		/// <returns> the raw data
		/// </returns>
		virtual public sbyte[] Data
		{
			get
			{
				return getRecord().Data;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The name</summary>
		private string name;
		
		/// <summary> The 0-based index in the name table</summary>
		private int index;
		
		/// <summary> Dummy indicators for overloading the constructor</summary>
		public class Biff7
		{
		}
		
		public static Biff7 biff7;
		
		// Constants which refer to the parse tokens after the string
		private const int cellReference = 0x3a;
		private const int areaReference = 0x3b;
		private const int subExpression = 0x29;
		private const int union = 0x10;
		
		/// <summary> A nested class to hold range information</summary>
		public class NameRange
		{
			private void  InitBlock(NameRecord enclosingInstance)
			{
				this.enclosingInstance = enclosingInstance;
			}
			private NameRecord enclosingInstance;
			/// <summary> Accessor for the first column
			/// 
			/// </summary>
			/// <returns> the index of the first column
			/// </returns>
			virtual public int FirstColumn
			{
				get
				{
					return columnFirst;
				}
				
			}
			/// <summary> Accessor for the first row
			/// 
			/// </summary>
			/// <returns> the index of the first row
			/// </returns>
			virtual public int FirstRow
			{
				get
				{
					return rowFirst;
				}
				
			}
			/// <summary> Accessor for the last column
			/// 
			/// </summary>
			/// <returns> the index of the last column
			/// </returns>
			virtual public int LastColumn
			{
				get
				{
					return columnLast;
				}
				
			}
			/// <summary> Accessor for the last row
			/// 
			/// </summary>
			/// <returns> the index of the last row
			/// </returns>
			virtual public int LastRow
			{
				get
				{
					return rowLast;
				}
				
			}
			/// <summary> Accessor for the first sheet
			/// 
			/// </summary>
			/// <returns>  the index of the external  sheet
			/// </returns>
			virtual public int ExternalSheet
			{
				get
				{
					return externalSheet;
				}
				
			}
			public NameRecord Enclosing_Instance
			{
				get
				{
					return enclosingInstance;
				}
				
			}
			/// <summary> The first column</summary>
			private int columnFirst;
			
			/// <summary> The first row</summary>
			private int rowFirst;
			
			/// <summary> The last column</summary>
			private int columnLast;
			
			/// <summary> The last row</summary>
			private int rowLast;
			
			/// <summary> The first sheet</summary>
			private int externalSheet;
			
			/// <summary> Constructor
			/// 
			/// </summary>
			/// <param name="s1">the sheet
			/// </param>
			/// <param name="c1">the first column
			/// </param>
			/// <param name="r1">the first row
			/// </param>
			/// <param name="c2">the last column
			/// </param>
			/// <param name="r2">the last row
			/// </param>
			internal NameRange(NameRecord enclosingInstance, int s1, int c1, int r1, int c2, int r2)
			{
				InitBlock(enclosingInstance);
				columnFirst = c1;
				rowFirst = r1;
				columnLast = c2;
				rowLast = r2;
				externalSheet = s1;
			}
		}
		
		/// <summary> The ranges referenced by this name</summary>
		private ArrayList ranges;
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <param name="ind">the index in the name table
		/// </param>
		internal NameRecord(Record t, WorkbookSettings ws, int ind):base(t)
		{
			index = ind;
			
			try
			{
				ranges = new ArrayList();
				
				sbyte[] data = getRecord().Data;
				int length = data[3];
				name = StringHelper.getString(data, length, 15, ws);
				int pos = length + 15;
				
				if (data[pos] == cellReference)
				{
					int sheet = IntegerHelper.getInt(data[pos + 1], data[pos + 2]);
					int row = IntegerHelper.getInt(data[pos + 3], data[pos + 4]);
					int columnMask = IntegerHelper.getInt(data[pos + 5], data[pos + 6]);
					int column = columnMask & 0xff;
					
					// Check that we are not dealing with offsets
					Assert.verify((columnMask & 0xc0000) == 0);
					
					NameRange r = new NameRange(this, sheet, column, row, column, row);
					ranges.Add(r);
				}
				else if (data[pos] == areaReference)
				{
					int sheet1 = 0;
					//        int sheet2 = 0;
					int r1 = 0;
					int columnMask = 0;
					int c1 = 0;
					int r2 = 0;
					int c2 = 0;
					NameRange range = null;
					
					while (pos < data.Length)
					{
						sheet1 = IntegerHelper.getInt(data[pos + 1], data[pos + 2]);
						r1 = IntegerHelper.getInt(data[pos + 3], data[pos + 4]);
						r2 = IntegerHelper.getInt(data[pos + 5], data[pos + 6]);
						
						columnMask = IntegerHelper.getInt(data[pos + 7], data[pos + 8]);
						c1 = columnMask & 0xff;
						
						// Check that we are not dealing with offsets
						Assert.verify((columnMask & 0xc0000) == 0);
						
						columnMask = IntegerHelper.getInt(data[pos + 9], data[pos + 10]);
						c2 = columnMask & 0xff;
						
						// Check that we are not dealing with offsets
						Assert.verify((columnMask & 0xc0000) == 0);
						
						range = new NameRange(this, sheet1, c1, r1, c2, r2);
						ranges.Add(range);
						
						pos += 11;
					}
				}
				else if (data[pos] == subExpression)
				{
					int sheet1 = 0;
					//        int sheet2 = 0;
					int r1 = 0;
					int columnMask = 0;
					int c1 = 0;
					int r2 = 0;
					int c2 = 0;
					NameRange range = null;
					
					// Consume unnecessary parsed tokens
					if (pos < data.Length && data[pos] != cellReference && data[pos] != areaReference)
					{
						if (data[pos] == subExpression)
						{
							pos += 3;
						}
						else if (data[pos] == union)
						{
							pos += 1;
						}
					}
					
					while (pos < data.Length)
					{
						sheet1 = IntegerHelper.getInt(data[pos + 1], data[pos + 2]);
						r1 = IntegerHelper.getInt(data[pos + 3], data[pos + 4]);
						r2 = IntegerHelper.getInt(data[pos + 5], data[pos + 6]);
						
						columnMask = IntegerHelper.getInt(data[pos + 7], data[pos + 8]);
						c1 = columnMask & 0xff;
						
						// Check that we are not dealing with offsets
						Assert.verify((columnMask & 0xc0000) == 0);
						
						columnMask = IntegerHelper.getInt(data[pos + 9], data[pos + 10]);
						c2 = columnMask & 0xff;
						
						// Check that we are not dealing with offsets
						Assert.verify((columnMask & 0xc0000) == 0);
						
						range = new NameRange(this, sheet1, c1, r1, c2, r2);
						ranges.Add(range);
						
						pos += 11;
						
						// Consume unnecessary parsed tokens
						if (pos < data.Length && data[pos] != cellReference && data[pos] != areaReference)
						{
							if (data[pos] == subExpression)
							{
								pos += 3;
							}
							else if (data[pos] == union)
							{
								pos += 1;
							}
						}
					}
				}
			}
			catch (System.Exception t1)
			{
				// Generate a warning
				// Names are really a nice to have, and we don't want to halt the
				// reading process for functionality that probably won't be used
				logger.warn("Cannot read name");
				name = "ERROR";
			}
		}
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <param name="ind">the index in the name table
		/// </param>
		/// <param name="dummy">dummy parameter to indicate a biff7 workbook
		/// </param>
		internal NameRecord(Record t, WorkbookSettings ws, int ind, Biff7 dummy):base(t)
		{
			index = ind;
			try
			{
				ranges = new ArrayList();
				sbyte[] data = getRecord().Data;
				int length = data[3];
				name = StringHelper.getString(data, length, 14, ws);
				
				int pos = length + 14;
				
				if (pos >= data.Length)
				{
					// There appears to be nothing after the name, so return
					return ;
				}
				
				if (data[pos] == cellReference)
				{
					int sheet = IntegerHelper.getInt(data[pos + 11], data[pos + 12]);
					int row = IntegerHelper.getInt(data[pos + 15], data[pos + 16]);
					int column = data[pos + 17];
					
					NameRange r = new NameRange(this, sheet, column, row, column, row);
					ranges.Add(r);
				}
				else if (data[pos] == areaReference)
				{
					int sheet1 = 0;
					int sheet2 = 0;
					int r1 = 0;
					//        int columnMask = 0;
					int c1 = 0;
					int r2 = 0;
					int c2 = 0;
					NameRange range = null;
					
					while (pos < data.Length)
					{
						sheet1 = IntegerHelper.getInt(data[pos + 11], data[pos + 12]);
						sheet2 = IntegerHelper.getInt(data[pos + 13], data[pos + 14]);
						r1 = IntegerHelper.getInt(data[pos + 15], data[pos + 16]);
						r2 = IntegerHelper.getInt(data[pos + 17], data[pos + 18]);
						
						c1 = data[pos + 19];
						c2 = data[pos + 20];
						
						range = new NameRange(this, sheet1, c1, r1, c2, r2);
						ranges.Add(range);
						
						pos += 21;
					}
				}
				else if (data[pos] == subExpression)
				{
					
					int sheet1 = 0;
					int sheet2 = 0;
					int r1 = 0;
					//        int columnMask = 0;
					int c1 = 0;
					int r2 = 0;
					int c2 = 0;
					NameRange range = null;
					
					// Consume unnecessary parsed tokens
					if (pos < data.Length && data[pos] != cellReference && data[pos] != areaReference)
					{
						if (data[pos] == subExpression)
						{
							pos += 3;
						}
						else if (data[pos] == union)
						{
							pos += 1;
						}
					}
					
					while (pos < data.Length)
					{
						sheet1 = IntegerHelper.getInt(data[pos + 11], data[pos + 12]);
						sheet2 = IntegerHelper.getInt(data[pos + 13], data[pos + 14]);
						r1 = IntegerHelper.getInt(data[pos + 15], data[pos + 16]);
						r2 = IntegerHelper.getInt(data[pos + 17], data[pos + 18]);
						
						c1 = data[pos + 19];
						c2 = data[pos + 20];
						
						range = new NameRange(this, sheet1, c1, r1, c2, r2);
						ranges.Add(range);
						
						pos += 21;
						
						// Consume unnecessary parsed tokens
						if (pos < data.Length && data[pos] != cellReference && data[pos] != areaReference)
						{
							if (data[pos] == subExpression)
							{
								pos += 3;
							}
							else if (data[pos] == union)
							{
								pos += 1;
							}
						}
					}
				}
			}
			catch (System.Exception t1)
			{
				// Generate a warning
				// Names are really a nice to have, and we don't want to halt the
				// reading process for functionality that probably won't be used
				logger.warn("Cannot read name.");
				name = "ERROR";
			}
		}
		static NameRecord()
		{
			logger = Logger.getLogger(typeof(NameRecord));
			biff7 = new Biff7();
		}
	}
}
