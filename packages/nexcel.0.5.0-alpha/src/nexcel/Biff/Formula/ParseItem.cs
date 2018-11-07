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
using NExcel;
namespace NExcel.Biff.Formula
{
	
	/// <summary> Abstract base class for an item in a formula parse tree</summary>
	abstract class ParseItem
	{
		/// <summary> Called by this class to initialize the parent</summary>
		virtual protected internal ParseItem Parent
		{
			set
			{
				parent = value;
			}
			
		}
		/// <summary> Gets the token representation of this item in RPN
		/// 
		/// </summary>
		/// <returns> the bytes applicable to this formula
		/// </returns>
		internal abstract sbyte[] Bytes{get;}
		/// <summary> The parent of this parse item</summary>
		private ParseItem parent;
		
		/// <summary> Volatile flag</summary>
		private bool volatileFunction;
		
		/// <summary> Indicates that the alternative token code should be used</summary>
		private bool alternateCode;
		
		/// <summary> Constructor</summary>
		public ParseItem()
		{
			volatileFunction = false;
			alternateCode = false;
		}
		
		/// <summary> Sets the volatile flag and ripples all the way up the parse tree</summary>
		protected internal virtual void  setVolatile()
		{
			volatileFunction = true;
			if (parent != null && !parent.isVolatile())
			{
				parent.setVolatile();
			}
		}
		
		/// <summary> Accessor for the volatile function
		/// 
		/// </summary>
		/// <returns> TRUE if the formula is volatile, FALSE otherwise
		/// </returns>
		internal bool isVolatile()
		{
			return volatileFunction;
		}
		
		/// <summary> Gets the string representation of this item</summary>
		/// <param name="ws">the workbook settings
		/// </param>
		public abstract void  getString(System.Text.StringBuilder buf);
		
		/// <summary> Adjusts all the relative cell references in this formula by the
		/// amount specified.  Used when copying formulas
		/// 
		/// </summary>
		/// <param name="colAdjust">the amount to add on to each relative cell reference
		/// </param>
		/// <param name="rowAdjust">the amount to add on to each relative row reference
		/// </param>
		public abstract void  adjustRelativeCellReferences(int colAdjust, int rowAdjust);
		
		/// <summary> Called when a column is inserted on the specified sheet.  Tells
		/// the formula  parser to update all of its cell references beyond this
		/// column
		/// 
		/// </summary>
		/// <param name="sheetIndex">the sheet on which the column was inserted
		/// </param>
		/// <param name="col">the column number which was inserted
		/// </param>
		/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
		/// column was inserted, FALSE otherwise
		/// </param>
		public abstract void  columnInserted(int sheetIndex, int col, bool currentSheet);
		
		/// <summary> Called when a column is inserted on the specified sheet.  Tells
		/// the formula  parser to update all of its cell references beyond this
		/// column
		/// 
		/// </summary>
		/// <param name="sheetIndex">the sheet on which the column was removed
		/// </param>
		/// <param name="col">the column number which was removed
		/// </param>
		/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
		/// column was inserted, FALSE otherwise
		/// </param>
		internal abstract void  columnRemoved(int sheetIndex, int col, bool currentSheet);
		
		/// <summary> Called when a column is inserted on the specified sheet.  Tells
		/// the formula  parser to update all of its cell references beyond this
		/// column
		/// 
		/// </summary>
		/// <param name="sheetIndex">the sheet on which the row was inserted
		/// </param>
		/// <param name="row">the row number which was inserted
		/// </param>
		/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
		/// column was inserted, FALSE otherwise
		/// </param>
		internal abstract void  rowInserted(int sheetIndex, int row, bool currentSheet);
		
		/// <summary> Called when a column is inserted on the specified sheet.  Tells
		/// the formula  parser to update all of its cell references beyond this
		/// column
		/// 
		/// </summary>
		/// <param name="sheetIndex">the sheet on which the row was removed
		/// </param>
		/// <param name="row">the row number which was removed
		/// </param>
		/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
		/// column was inserted, FALSE otherwise
		/// </param>
		internal abstract void  rowRemoved(int sheetIndex, int row, bool currentSheet);
		
		/// <summary> Tells the operands to use the alternate code</summary>
		protected internal virtual void  setAlternateCode()
		{
			alternateCode = true;
		}
		
		
		/// <summary> Accessor for the alternate code flag
		/// 
		/// </summary>
		/// <returns> TRUE to use the alternate code, FALSE otherwise
		/// </returns>
		protected internal bool useAlternateCode()
		{
			return alternateCode;
		}
	}
}
