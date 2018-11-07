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
using NExcel.Biff.Formula;
namespace NExcel.Biff
{
	
	/// <summary> A helper to transform between excel cell references and
	/// sheet:column:row notation
	/// Because this function will be called when generating a string
	/// representation of a formula, the cell reference will merely
	/// be appened to the string buffer instead of returning a full
	/// blooded string, for performance reasons
	/// </summary>
	public sealed class CellReferenceHelper
	{
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The character which indicates whether a reference is fixed</summary>
		private const char fixedInd = '$';
		
		/// <summary> Constructor to prevent instantiation</summary>
		private CellReferenceHelper()
		{
		}
		
		/// <summary> Gets the cell reference 
		/// 
		/// </summary>
		/// <param name="">column
		/// </param>
		/// <param name="">row
		/// </param>
		/// <param name="">buf
		/// </param>
		public static void  getCellReference(int column, int row, System.Text.StringBuilder buf)
		{
			// Put the column letter into the buffer
			getColumnReference(column, buf);
			
			// Add the row into the buffer
			buf.Append(System.Convert.ToString(row + 1));
		}
		
		/// <summary> Overloaded method which prepends $ for absolute reference
		/// 
		/// </summary>
		/// <param name="">column
		/// </param>
		/// <param name="colabs">TRUE if the column reference is absolute
		/// </param>
		/// <param name="">row
		/// </param>
		/// <param name="rowabs">TRUE if the row reference is absolute
		/// </param>
		/// <param name="">buf
		/// </param>
		public static void  getCellReference(int column, bool colabs, int row, bool rowabs, System.Text.StringBuilder buf)
		{
			if (colabs)
			{
				buf.Append(fixedInd);
			}
			
			// Put the column letter into the buffer
			getColumnReference(column, buf);
			
			if (rowabs)
			{
				buf.Append(fixedInd);
			}
			
			// Add the row into the buffer
			buf.Append(System.Convert.ToString(row + 1));
		}
		
		/// <summary> Gets the column letter corresponding to the 0-based column number
		/// 
		/// </summary>
		/// <param name="column">the column number
		/// </param>
		/// <returns> the letter for that column number
		/// </returns>
		public static string getColumnReference(int column)
		{
			System.Text.StringBuilder buf = new System.Text.StringBuilder();
			getColumnReference(column, buf);
			return buf.ToString();
		}
		
		/// <summary> Gets the column letter corresponding to the 0-based column number
		/// 
		/// </summary>
		/// <param name="column">the column number
		/// </param>
		/// <param name="buf">the string buffer in which to write the column letter
		/// </param>
		public static void  getColumnReference(int column, System.Text.StringBuilder buf)
		{
			int v = column / 26;
			int r = column % 26;
			
			System.Text.StringBuilder tmp = new System.Text.StringBuilder();
			while (v != 0)
			{
				char col = (char) ((int) 'A' + r);
				
				tmp.Append(col);
				
				r = v % 26 - 1; // subtract one because only rows >26 preceded by A
				v = v / 26;
			}
			
			char col2 = (char) ((int) 'A' + r);
			tmp.Append(col2);
			
			// Insert into the proper string buffer in reverse order
			for (int i = tmp.Length - 1; i >= 0; i--)
			{
				buf.Append(tmp[i]);
			}
		}
		
		/// <summary> Gets the fully qualified cell reference given the column, row
		/// external sheet reference etc
		/// 
		/// </summary>
		/// <param name="">sheet
		/// </param>
		/// <param name="">column
		/// </param>
		/// <param name="">row
		/// </param>
		/// <param name="">workbook
		/// </param>
		/// <param name="">buf
		/// </param>
		public static void  getCellReference(int sheet, int column, int row, ExternalSheet workbook, System.Text.StringBuilder buf)
		{
			buf.Append('\'');
			buf.Append(workbook.getExternalSheetName(sheet));
			buf.Append('\'');
			buf.Append('!');
			getCellReference(column, row, buf);
		}
		
		/// <summary> Gets the fully qualified cell reference given the column, row
		/// external sheet reference etc
		/// 
		/// </summary>
		/// <param name="">sheet
		/// </param>
		/// <param name="">column
		/// </param>
		/// <param name="colabs">TRUE if the column is an absolute reference
		/// </param>
		/// <param name="">row
		/// </param>
		/// <param name="rowabs">TRUE if the row is an absolute reference
		/// </param>
		/// <param name="">workbook
		/// </param>
		/// <param name="">buf
		/// </param>
		public static void  getCellReference(int sheet, int column, bool colabs, int row, bool rowabs, ExternalSheet workbook, System.Text.StringBuilder buf)
		{
			buf.Append('\'');
			buf.Append(workbook.getExternalSheetName(sheet));
			buf.Append('\'');
			buf.Append('!');
			getCellReference(column, colabs, row, rowabs, buf);
		}
		
		/// <summary> Gets the fully qualified cell reference given the column, row
		/// external sheet reference etc
		/// 
		/// </summary>
		/// <param name="">sheet
		/// </param>
		/// <param name="">column
		/// </param>
		/// <param name="">row
		/// </param>
		/// <param name="">workbook
		/// </param>
		/// <returns> the cell reference in the form 'Sheet 1'!A1
		/// </returns>
		public static string getCellReference(int sheet, int column, int row, ExternalSheet workbook)
		{
			System.Text.StringBuilder sb = new System.Text.StringBuilder();
			getCellReference(sheet, column, row, workbook, sb);
			return sb.ToString();
		}
		
		
		/// <summary> Gets the cell reference for the specified column and row
		/// 
		/// </summary>
		/// <param name="">column
		/// </param>
		/// <param name="">row
		/// </param>
		/// <returns>
		/// </returns>
		public static string getCellReference(int column, int row)
		{
			System.Text.StringBuilder buf = new System.Text.StringBuilder();
			getCellReference(column, row, buf);
			return buf.ToString();
		}
		
		/// <summary> Gets the columnn number of the string cell reference
		/// 
		/// </summary>
		/// <param name="s">the string to parse
		/// </param>
		/// <returns> the column portion of the cell reference
		/// </returns>
		public static int getColumn(string s)
		{
			int colnum = 0;
			int numindex = getNumberIndex(s);
			
			string s2 = s.ToUpper();
			
			int startPos = 0;
			if (s[0] == fixedInd)
			{
				startPos = 1;
			}
			
			int endPos = numindex;
			if (s[numindex - 1] == fixedInd)
			{
				endPos--;
			}
			
			for (int i = startPos; i < endPos; i++)
			{
				
				if (i != startPos)
				{
					colnum = (colnum + 1) * 26;
				}
				colnum += (int) s2[i] - (int) 'A';
			}
			
			return colnum;
		}
		
		/// <summary> Gets the row number of the cell reference</summary>
		public static int getRow(string s)
		{
			try
			{
				return (System.Int32.Parse(s.Substring(getNumberIndex(s))) - 1);
			}
			catch (System.FormatException e)
			{
				logger.warn(e, e);
				return 0xffff;
			}
		}
		
		/// <summary> Finds the position where the first number occurs in the string</summary>
		private static int getNumberIndex(string s)
		{
			// Find the position of the first number
			bool numberFound = false;
			int pos = 0;
			char c = '\x0000';
			
			while (!numberFound && pos < s.Length)
			{
				c = s[pos];
				
				if (c >= '0' && c <= '9')
				{
					numberFound = true;
				}
				else
				{
					pos++;
				}
			}
			
			return pos;
		}
		
		/// <summary> Sees if the column component is relative or not
		/// 
		/// </summary>
		/// <param name="">s
		/// </param>
		/// <returns> TRUE if the column is relative, FALSE otherwise
		/// </returns>
		public static bool isColumnRelative(string s)
		{
			return s[0] != fixedInd;
		}
		
		/// <summary> Sees if the row component is relative or not
		/// 
		/// </summary>
		/// <param name="">s
		/// </param>
		/// <returns> TRUE if the row is relative, FALSE otherwise
		/// </returns>
		public static bool isRowRelative(string s)
		{
			return s[getNumberIndex(s) - 1] != fixedInd;
		}
		static CellReferenceHelper()
		{
			logger = Logger.getLogger(typeof(CellReferenceHelper));
		}
	}
}
