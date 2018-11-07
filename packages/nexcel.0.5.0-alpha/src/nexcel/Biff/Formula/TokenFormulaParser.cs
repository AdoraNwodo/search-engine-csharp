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
	
	/// <summary> Parses the excel ptgs into a parse tree</summary>
	class TokenFormulaParser : Parser
	{
		/// <summary> Gets the formula as a string</summary>
		virtual public string Formula
		{
			get
			{
				System.Text.StringBuilder sb = new System.Text.StringBuilder();
				root.getString(sb);
				return sb.ToString();
			}
			
		}
		/// <summary> Gets the bytes for the formula. This takes into account any
		/// token mapping necessary because of shared formulas
		/// 
		/// </summary>
		/// <returns> the bytes in RPN
		/// </returns>
		virtual public sbyte[] Bytes
		{
			get
			{
				return root.Bytes;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The Excel ptgs</summary>
		private sbyte[] tokenData;
		
		/// <summary> The cell containing the formula.  This is used in order to determine
		/// relative cell values
		/// </summary>
		private Cell relativeTo;
		
		/// <summary> The current position within the array</summary>
		private int pos;
		
		/// <summary> The parse tree</summary>
		private ParseItem root;
		
		/// <summary> The hash table of items that have been parsed</summary>
		private Stack tokenStack;
		
		/// <summary> A reference to the workbook which holds the external sheet
		/// information
		/// </summary>
		private ExternalSheet workbook;
		
		/// <summary> A reference to the name table</summary>
		private WorkbookMethods nameTable;
		
		/// <summary> The workbook settings</summary>
		private WorkbookSettings settings;
		
		/// <summary> Constructor</summary>
		public TokenFormulaParser(sbyte[] data, Cell c, ExternalSheet es, WorkbookMethods nt, WorkbookSettings ws)
		{
			tokenData = data;
			pos = 0;
			relativeTo = c;
			workbook = es;
			nameTable = nt;
			tokenStack = new Stack();
			settings = ws;
		}
		
		/// <summary> Parses the sublist of tokens.  In most cases this will equate to
		/// the full list
		/// 
		/// </summary>
		/// <exception cref=""> FormulaException
		/// </exception>
		public virtual void  parse()
		{
			parseSubExpression(tokenData.Length);
			
			// Finally, there should be one thing left on the stack.  Get that
			// and add it to the root node
			root = (ParseItem) tokenStack.Pop();
			
			Assert.verify(tokenStack.Count <= 0);
		}
		
		/// <summary> Parses the sublist of tokens.  In most cases this will equate to
		/// the full list
		/// 
		/// </summary>
		/// <param name="len">the .Length of the subexpression to parse
		/// </param>
		/// <exception cref=""> FormulaException
		/// </exception>
		private void  parseSubExpression(int len)
		{
			int tokenVal = 0;
			Token t = null;
			
			// Indicates that we are parsing the incredibly complicated and
			// hacky if construct that MS saw fit to include, the gits
			Stack ifStack = new Stack();
			
			// The end position of the sub-expression
			int endpos = pos + len;
			
			while (pos < endpos)
			{
				tokenVal = tokenData[pos];
				pos++;
				
				t = Token.getToken(tokenVal);
				
				if (t == Token.UNKNOWN)
				{
					throw new FormulaException(FormulaException.unrecognizedToken, tokenVal);
				}
				
				Assert.verify(t != Token.UNKNOWN);
				
				// Operands
				if (t == Token.REF)
				{
					CellReference cr = new CellReference(relativeTo);
					pos += cr.read(tokenData, pos);
					tokenStack.Push(cr);
				}
				else if (t == Token.REFV)
				{
					SharedFormulaCellReference cr = new SharedFormulaCellReference(relativeTo);
					pos += cr.read(tokenData, pos);
					tokenStack.Push(cr);
				}
				else if (t == Token.REF3D)
				{
					CellReference3d cr = new CellReference3d(relativeTo, workbook);
					pos += cr.read(tokenData, pos);
					tokenStack.Push(cr);
				}
				else if (t == Token.AREA)
				{
					Area a = new Area();
					pos += a.read(tokenData, pos);
					tokenStack.Push(a);
				}
				else if (t == Token.AREAV)
				{
					SharedFormulaArea a = new SharedFormulaArea(relativeTo);
					pos += a.read(tokenData, pos);
					tokenStack.Push(a);
				}
				else if (t == Token.AREA3D)
				{
					Area3d a = new Area3d(workbook);
					pos += a.read(tokenData, pos);
					tokenStack.Push(a);
				}
				else if (t == Token.NAME)
				{
					Name n = new Name();
					pos += n.read(tokenData, pos);
					tokenStack.Push(n);
				}
				else if (t == Token.NAMED_RANGE)
				{
					NameRange nr = new NameRange(nameTable);
					pos += nr.read(tokenData, pos);
					tokenStack.Push(nr);
				}
				else if (t == Token.INTEGER)
				{
					IntegerValue i = new IntegerValue();
					pos += i.read(tokenData, pos);
					tokenStack.Push(i);
				}
				else if (t == Token.DOUBLE)
				{
					DoubleValue d = new DoubleValue();
					pos += d.read(tokenData, pos);
					tokenStack.Push(d);
				}
				else if (t == Token.BOOL)
				{
					BooleanValue bv = new BooleanValue();
					pos += bv.read(tokenData, pos);
					tokenStack.Push(bv);
				}
				else if (t == Token.STRING)
				{
					StringValue sv = new StringValue(settings);
					pos += sv.read(tokenData, pos);
					tokenStack.Push(sv);
				}
				else if (t == Token.MISSING_ARG)
				{
					MissingArg ma = new MissingArg();
					pos += ma.read(tokenData, pos);
					tokenStack.Push(ma);
				}
				// Unary Operators
				else if (t == Token.UNARY_PLUS)
				{
					UnaryPlus up = new UnaryPlus();
					pos += up.read(tokenData, pos);
					addOperator(up);
				}
				else if (t == Token.UNARY_MINUS)
				{
					UnaryMinus um = new UnaryMinus();
					pos += um.read(tokenData, pos);
					addOperator(um);
				}
				else if (t == Token.PERCENT)
				{
					Percent p = new Percent();
					pos += p.read(tokenData, pos);
					addOperator(p);
				}
				// Binary Operators
				else if (t == Token.SUBTRACT)
				{
					Subtract s = new Subtract();
					pos += s.read(tokenData, pos);
					addOperator(s);
				}
				else if (t == Token.ADD)
				{
					Add s = new Add();
					pos += s.read(tokenData, pos);
					addOperator(s);
				}
				else if (t == Token.MULTIPLY)
				{
					Multiply s = new Multiply();
					pos += s.read(tokenData, pos);
					addOperator(s);
				}
				else if (t == Token.DIVIDE)
				{
					Divide s = new Divide();
					pos += s.read(tokenData, pos);
					addOperator(s);
				}
				else if (t == Token.CONCAT)
				{
					Concatenate c = new Concatenate();
					pos += c.read(tokenData, pos);
					addOperator(c);
				}
				else if (t == Token.POWER)
				{
					Power p = new Power();
					pos += p.read(tokenData, pos);
					addOperator(p);
				}
				else if (t == Token.LESS_THAN)
				{
					LessThan lt = new LessThan();
					pos += lt.read(tokenData, pos);
					addOperator(lt);
				}
				else if (t == Token.LESS_EQUAL)
				{
					LessEqual lte = new LessEqual();
					pos += lte.read(tokenData, pos);
					addOperator(lte);
				}
				else if (t == Token.GREATER_THAN)
				{
					GreaterThan gt = new GreaterThan();
					pos += gt.read(tokenData, pos);
					addOperator(gt);
				}
				else if (t == Token.GREATER_EQUAL)
				{
					GreaterEqual gte = new GreaterEqual();
					pos += gte.read(tokenData, pos);
					addOperator(gte);
				}
				else if (t == Token.NOT_EQUAL)
				{
					NotEqual ne = new NotEqual();
					pos += ne.read(tokenData, pos);
					addOperator(ne);
				}
				else if (t == Token.EQUAL)
				{
					Equal e = new Equal();
					pos += e.read(tokenData, pos);
					addOperator(e);
				}
				else if (t == Token.PARENTHESIS)
				{
					Parenthesis p = new Parenthesis();
					pos += p.read(tokenData, pos);
					addOperator(p);
				}
				// Functions
				else if (t == Token.ATTRIBUTE)
				{
					Attribute a = new Attribute(settings);
					pos += a.read(tokenData, pos);
					
					if (a.Sum)
					{
						addOperator(a);
					}
					else if (a.isIf())
					{
						// Add it to a special stack for ifs
						ifStack.Push(a);
					}
				}
				else if (t == Token.FUNCTION)
				{
					BuiltInFunction bif = new BuiltInFunction(settings);
					pos += bif.read(tokenData, pos);
					
					addOperator(bif);
				}
				else if (t == Token.FUNCTIONVARARG)
				{
					VariableArgFunction vaf = new VariableArgFunction(settings);
					pos += vaf.read(tokenData, pos);
					
					if (vaf.Function != Function.ATTRIBUTE)
					{
						addOperator(vaf);
					}
					else
					{
						// This is part of an IF function.  Get the operands, but then
						// add it to the top of the if stack
						vaf.getOperands(tokenStack);
						
						Attribute ifattr = null;
						if (ifStack.Count <= 0)
						{
							ifattr = new Attribute(settings);
						}
						else
						{
							ifattr = (Attribute) ifStack.Pop();
						}
						
						ifattr.IfConditions = vaf;
						tokenStack.Push(ifattr);
					}
				}
				// Other things
				else if (t == Token.MEM_FUNC)
				{
					MemFunc memFunc = new MemFunc();
					pos += memFunc.read(tokenData, pos);
					
					// Create new tokenStack for the sub expression
					Stack oldStack = tokenStack;
					tokenStack = new Stack();
					
					parseSubExpression(memFunc.Length);
					
					ParseItem[] subexpr = new ParseItem[tokenStack.Count];
					int i = 0;
					while (tokenStack.Count > 0)
					{
						subexpr[i] = (ParseItem) tokenStack.Pop();
						i++;
					}
					
					memFunc.SubExpression = subexpr;
					
					tokenStack = oldStack;
					tokenStack.Push(memFunc);
				}
			}
		}
		
		/// <summary> Adds the specified operator to the parse tree, taking operands off
		/// the stack as appropriate
		/// </summary>
		private void  addOperator(Operator o)
		{
			// Get the operands off the stack
			o.getOperands(tokenStack);
			
			// Add this operator onto the stack
			tokenStack.Push(o);
		}
		
		/// <summary> Adjusts all the relative cell references in this formula by the
		/// amount specified.  Used when copying formulas
		/// 
		/// </summary>
		/// <param name="colAdjust">the amount to add on to each relative cell reference
		/// </param>
		/// <param name="rowAdjust">the amount to add on to each relative row reference
		/// </param>
		public virtual void  adjustRelativeCellReferences(int colAdjust, int rowAdjust)
		{
			root.adjustRelativeCellReferences(colAdjust, rowAdjust);
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
		public virtual void  columnInserted(int sheetIndex, int col, bool currentSheet)
		{
			root.columnInserted(sheetIndex, col, currentSheet);
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
		public virtual void  columnRemoved(int sheetIndex, int col, bool currentSheet)
		{
			root.columnRemoved(sheetIndex, col, currentSheet);
		}
		
		/// <summary> Called when a column is inserted on the specified sheet.  Tells
		/// the formula  parser to update all of its cell references beyond this
		/// column
		/// 
		/// </summary>
		/// <param name="sheetIndex">the sheet on which the column was inserted
		/// </param>
		/// <param name="row">the column number which was inserted
		/// </param>
		/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
		/// column was inserted, FALSE otherwise
		/// </param>
		public virtual void  rowInserted(int sheetIndex, int row, bool currentSheet)
		{
			root.rowInserted(sheetIndex, row, currentSheet);
		}
		
		/// <summary> Called when a column is inserted on the specified sheet.  Tells
		/// the formula  parser to update all of its cell references beyond this
		/// column
		/// 
		/// </summary>
		/// <param name="sheetIndex">the sheet on which the column was removed
		/// </param>
		/// <param name="row">the column number which was removed
		/// </param>
		/// <param name="currentSheet">TRUE if this formula is on the sheet in which the
		/// column was inserted, FALSE otherwise
		/// </param>
		public virtual void  rowRemoved(int sheetIndex, int row, bool currentSheet)
		{
			root.rowRemoved(sheetIndex, row, currentSheet);
		}
		static TokenFormulaParser()
		{
			logger = Logger.getLogger(typeof(TokenFormulaParser));
		}
	}
}
