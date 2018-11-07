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
namespace NExcel.Biff.Formula
{
	/// <summary> A special attribute control token - typically either a SUM function
	/// or an IF function
	/// </summary>
	
	class Attribute:Operator, ParsedThing
	{
		/// <summary> Sets the if conditions for this attribute, if it represents an IF function
		/// 
		/// </summary>
		/// <param name="vaf">a <code>VariableArgFunction</code> value
		/// </param>
		virtual internal VariableArgFunction IfConditions
		{
			set
			{
				ifConditions = value;
				
				// Sometimes there is not Attribute token, so we need to create
				// an attribute out of thin air.  In that case, make sure the if mask
				options |= ifMask;
			}
			
		}
		/// <summary> Queries whether this attribute is a function
		/// 
		/// </summary>
		/// <returns> TRUE if this is a function, FALSE otherwise
		/// </returns>
		virtual public bool Function
		{
			get
			{
				return (options & (sumMask | ifMask)) != 0;
			}
			
		}
		/// <summary> Queries whether this attribute is a sum
		/// 
		/// </summary>
		/// <returns> TRUE if this is SUM, FALSE otherwise
		/// </returns>
		virtual public bool Sum
		{
			get
			{
				return (options & sumMask) != 0;
			}
			
		}
		/// <summary> Queries whether this attribute is a goto
		/// 
		/// </summary>
		/// <returns> TRUE if this is a goto, FALSE otherwise
		/// </returns>
		virtual public bool Goto
		{
			get
			{
				return (options & gotoMask) != 0;
			}
			
		}
		/// <summary> Gets the token representation of this item in RPN.  The Attribute
		/// token is a special case, which overrides anything useful we could do
		/// in the base class
		/// 
		/// </summary>
		/// <returns> the bytes applicable to this formula
		/// </returns>
		override internal sbyte[] Bytes
		{
			get
			{
				sbyte[] data = new sbyte[0];
				if (Sum)
				{
					// Get the data for the operands
					ParseItem[] operands = getOperands();
					
					// Get the operands in reverse order to get the RPN
					for (int i = operands.Length - 1; i >= 0; i--)
					{
						sbyte[] opdata = operands[i].Bytes;
						
						// Grow the array
						sbyte[] cnewdata = new sbyte[data.Length + opdata.Length];
						Array.Copy(data, 0, cnewdata, 0, data.Length);
						Array.Copy(opdata, 0, cnewdata, data.Length, opdata.Length);
						data = cnewdata;
					}
					
					// Add on the operator byte
					sbyte[] newdata = new sbyte[data.Length + 4];
					Array.Copy(data, 0, newdata, 0, data.Length);
					newdata[data.Length] = Token.ATTRIBUTE.Code;
					newdata[data.Length + 1] = sumMask;
					data = newdata;
				}
				else if (isIf())
				{
					return getIf();
				}
				
				return data;
			}
			
		}
		/// <summary> Gets the precedence for this operator.  Operator precedents run from 
		/// 1 to 5, one being the highest, 5 being the lowest
		/// 
		/// </summary>
		/// <returns> the operator precedence
		/// </returns>
		override internal int Precedence
		{
			get
			{
				return 3;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The options used by the attribute</summary>
		private int options;
		
		/// <summary> The word contained in this attribute</summary>
		private int word;
		
		/// <summary> The workbook settings</summary>
		private WorkbookSettings settings;
		
		private const int sumMask = 0x10;
		private const int ifMask = 0x02;
		private const int gotoMask = 0x08;
		
		/// <summary> If this attribute is an IF functions, sets the associated if conditions</summary>
		private VariableArgFunction ifConditions;
		
		/// <summary> Constructor</summary>
		public Attribute(WorkbookSettings ws)
		{
			settings = ws;
		}
		
		/// <summary> Constructor for use when this is called when parsing a string
		/// 
		/// </summary>
		/// <param name="sf">the built in function
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public Attribute(StringFunction sf, WorkbookSettings ws)
		{
			settings = ws;
			
			if (sf.getFunction(settings) == NExcel.Biff.Formula.Function.SUM)
			{
				options |= sumMask;
			}
			else if (sf.getFunction(settings) == NExcel.Biff.Formula.Function.IF)
			{
				options |= ifMask;
			}
		}
		
		/// <summary> Reads the ptg data from the array starting at the specified position
		/// 
		/// </summary>
		/// <param name="data">the RPN array
		/// </param>
		/// <param name="pos">the current position in the array, excluding the ptg identifier
		/// </param>
		/// <returns> the number of bytes read
		/// </returns>
		public virtual int read(sbyte[] data, int pos)
		{
			options = data[pos];
			word = IntegerHelper.getInt(data[pos + 1], data[pos + 2]);
			return 3;
		}
		
		/// <summary> Queries whether this attribute is an IF
		/// 
		/// </summary>
		/// <returns> TRUE if this is an IF, FALSE otherwise
		/// </returns>
		public virtual bool isIf()
		{
			return (options & ifMask) != 0;
		}
		
		/// <summary> Gets the operands for this operator from the stack</summary>
		/// <summary> Gets the operands for this operator from the stack</summary>
		public override void getOperands(Stack s)
		{
		if ((options & sumMask) != 0)
		{
		ParseItem o1 = (ParseItem) s.Pop();
		
		add(o1);
		}
		else if ((options & ifMask) != 0)
		{
		ParseItem o1 = (ParseItem) s.Pop();
		add(o1);
		}
		}
		
		
		public override void  getString(System.Text.StringBuilder buf)
		{
			if ((options & sumMask) != 0)
			{
				ParseItem[] operands = getOperands();
				buf.Append(NExcel.Biff.Formula.Function.SUM.getName(settings));
				buf.Append('(');
				operands[0].getString(buf);
				buf.Append(')');
			}
			else if ((options & ifMask) != 0)
			{
				buf.Append(NExcel.Biff.Formula.Function.IF.getName(settings));
				buf.Append('(');
				
				ParseItem[] operands = ifConditions.getOperands();
				
				// Operands are in the correct order for IFs
				for (int i = 0; i < operands.Length; i++)
				{
					operands[i].getString(buf);
					buf.Append(',');
				}
				operands[0].getString(buf);
				buf.Append(')');
			}
		}
		
		/// <summary> Gets the associated if conditions with this attribute
		/// 
		/// </summary>
		/// <returns> the associated if conditions
		/// </returns>
		private sbyte[] getIf()
		{
			ParseItem[] operands = ifConditions.getOperands();
			
			// The position of the offset to the false portion of the expression
			int falseOffsetPos = 0;
			int gotoEndPos = 0;
			int numArgs = operands.Length;
			
			// First, write out the conditions
			sbyte[] data = operands[0].Bytes;
			
			// Grow the array by three and write out the optimized if attribute
			int pos = data.Length;
			sbyte[] newdata = new sbyte[data.Length + 4];
			Array.Copy(data, 0, newdata, 0, data.Length);
			data = newdata;
			data[pos] = Token.ATTRIBUTE.Code;
			data[pos + 1] = (sbyte) (0x2);
			falseOffsetPos = pos + 2;
			
			// Get the true portion of the expression and add it to the array
			sbyte[] truedata = operands[1].Bytes;
			newdata = new sbyte[data.Length + truedata.Length];
			Array.Copy(data, 0, newdata, 0, data.Length);
			Array.Copy(truedata, 0, newdata, data.Length, truedata.Length);
			data = newdata;
			
			// Grow the array by three and write out the goto end attribute
			pos = data.Length;
			newdata = new sbyte[data.Length + 4];
			Array.Copy(data, 0, newdata, 0, data.Length);
			data = newdata;
			data[pos] = Token.ATTRIBUTE.Code;
			data[pos + 1] = (sbyte) (0x8);
			gotoEndPos = pos + 2;
			
			// If the false condition exists, then add that to the array
			if (numArgs > 2)
			{
				// Set the offset to the false expression to be the current position
				IntegerHelper.getTwoBytes(data.Length - falseOffsetPos - 2, data, falseOffsetPos);
				
				// Copy in the false expression
				sbyte[] falsedata = operands[numArgs - 1].Bytes;
				newdata = new sbyte[data.Length + falsedata.Length];
				Array.Copy(data, 0, newdata, 0, data.Length);
				Array.Copy(falsedata, 0, newdata, data.Length, falsedata.Length);
				data = newdata;
				
				// Write the goto to skip over the varargs token
				pos = data.Length;
				newdata = new sbyte[data.Length + 4];
				Array.Copy(data, 0, newdata, 0, data.Length);
				data = newdata;
				data[pos] = Token.ATTRIBUTE.Code;
				data[pos + 1] = (sbyte) (0x8);
				data[pos + 2] = (sbyte) (0x3);
			}
			
			// Grow the array and write out the varargs function
			pos = data.Length;
			newdata = new sbyte[data.Length + 4];
			Array.Copy(data, 0, newdata, 0, data.Length);
			data = newdata;
			data[pos] = Token.FUNCTIONVARARG.Code;
			data[pos + 1] = (sbyte) numArgs;
			data[pos + 2] = 1;
			data[pos + 3] = 0; // indicates the end of the expression
			
			// Position the final offsets
			int endPos = data.Length - 1;
			
			if (numArgs < 3)
			{
				// Set the offset to the false expression to be the current position
				IntegerHelper.getTwoBytes(endPos - falseOffsetPos - 5, data, falseOffsetPos);
			}
			
			// Set the offset after the true expression
			IntegerHelper.getTwoBytes(endPos - gotoEndPos - 2, data, gotoEndPos);
			
			return data;
		}
		
		/// <summary> Default behaviour is to do nothing
		/// 
		/// </summary>
		/// <param name="colAdjust">the amount to add on to each relative cell reference
		/// </param>
		/// <param name="rowAdjust">the amount to add on to each relative row reference
		/// </param>
		public override void  adjustRelativeCellReferences(int colAdjust, int rowAdjust)
		{
			ParseItem[] operands = getOperands();
			for (int i = 0; i < operands.Length; i++)
			{
				operands[i].adjustRelativeCellReferences(colAdjust, rowAdjust);
			}
		}
		
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
		public override void  columnInserted(int sheetIndex, int col, bool currentSheet)
		{
			ParseItem[] operands = getOperands();
			for (int i = 0; i < operands.Length; i++)
			{
				operands[i].columnInserted(sheetIndex, col, currentSheet);
			}
		}
		
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
		internal override void  columnRemoved(int sheetIndex, int col, bool currentSheet)
		{
			ParseItem[] operands = getOperands();
			for (int i = 0; i < operands.Length; i++)
			{
				operands[i].columnRemoved(sheetIndex, col, currentSheet);
			}
		}
		
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
		internal override void  rowInserted(int sheetIndex, int row, bool currentSheet)
		{
			ParseItem[] operands = getOperands();
			for (int i = 0; i < operands.Length; i++)
			{
				operands[i].rowInserted(sheetIndex, row, currentSheet);
			}
		}
		
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
		internal override void  rowRemoved(int sheetIndex, int row, bool currentSheet)
		{
			ParseItem[] operands = getOperands();
			for (int i = 0; i < operands.Length; i++)
			{
				operands[i].rowRemoved(sheetIndex, row, currentSheet);
			}
		}
		static Attribute()
		{
			logger = Logger.getLogger(typeof(Attribute));
		}
	}
}
