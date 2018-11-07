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
	
	/// <summary> Parses a string formula into a parse tree</summary>
	class StringFormulaParser : Parser
	{
		/// <summary> Gets the list of lexical tokens using the generated lexical analyzer
		/// 
		/// </summary>
		/// <returns> the list of tokens
		/// </returns>
		/// <exception cref=""> FormulaException if an error occurs
		/// </exception>
		private ArrayList Tokens
		{
			get
			{
				ArrayList tokens = new ArrayList();
				
				Yylex lex = new Yylex(formula);
				lex.ExternalSheet = externalSheet;
				lex.NameTable = nameTable;
				try
				{
					ParseItem pi = lex.yylex();
					while (pi != null)
					{
						tokens.Add(pi);
						pi = lex.yylex();
					}
				}
				catch (System.IO.IOException e)
				{
					logger.warn(e.Message);
				}
				catch (System.ApplicationException e)
				{
					throw new FormulaException(FormulaException.lexicalError, formula + " at char  " + lex.Pos);
				}
				
				return tokens;
			}
			
		}
		/// <summary> Gets the formula as a string.  Uses the parse tree to do this, and
		/// does not simply return whatever string was passed in
		/// </summary>
		virtual public string Formula
		{
			get
			{
				if ((System.Object) parsedFormula == null)
				{
					System.Text.StringBuilder sb = new System.Text.StringBuilder();
					root.getString(sb);
					parsedFormula = sb.ToString();
				}
				
				return parsedFormula;
			}
			
		}
		/// <summary> Gets the bytes for the formula
		/// 
		/// </summary>
		/// <returns> the bytes in RPN
		/// </returns>
		virtual public sbyte[] Bytes
		{
			get
			{
				sbyte[] bytes = root.Bytes;
				
				if (root.isVolatile())
				{
					sbyte[] newBytes = new sbyte[bytes.Length + 4];
					Array.Copy(bytes, 0, newBytes, 4, bytes.Length);
					newBytes[0] = Token.ATTRIBUTE.Code;
					newBytes[1] = (sbyte) 0x1;
					bytes = newBytes;
				}
				
				return bytes;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The formula string passed to this object</summary>
		private string formula;
		
		/// <summary> The parsed formula string, as retrieved from the parse tree</summary>
		private string parsedFormula;
		
		/// <summary> The parse tree</summary>
		private ParseItem root;
		
		/// <summary> The stack argument used when parsing a function in order to
		/// pass multiple arguments back to the calling method
		/// </summary>
		private Stack arguments;
		
		/// <summary> The workbook settings</summary>
		private WorkbookSettings settings;
		
		/// <summary> A handle to the external sheet</summary>
		private ExternalSheet externalSheet;
		
		/// <summary> A handle to the name table</summary>
		private WorkbookMethods nameTable;
		
		/// <summary> Constructor</summary>
		/// <param name="">f
		/// </param>
		/// <param name="">ws
		/// </param>
		public StringFormulaParser(string f, ExternalSheet es, WorkbookMethods nt, WorkbookSettings ws)
		{
			formula = f;
			settings = ws;
			externalSheet = es;
			nameTable = nt;
		}
		
		/// <summary> Parses the list of tokens
		/// 
		/// </summary>
		/// <exception cref=""> FormulaException
		/// </exception>
		public virtual void  parse()
		{
			ArrayList tokens = Tokens;
			
			root = parseCurrent(tokens);
		}
		
		/// <summary> Recursively parses the token array.  Recursion is used in order
		/// to evaluate parentheses and function arguments
		/// 
		/// </summary>
		/// <param name="i">an iterator of tokens
		/// </param>
		/// <returns> the root node of the current parse stack
		/// </returns>
		/// <exception cref=""> FormulaException if an error occurs
		/// </exception>
		private ParseItem parseCurrent(ArrayList pis)
		{
			Stack stack = new Stack();
			Stack operators = new Stack();
			Stack args = null; // we usually don't need this
			
			bool parenthesesClosed = false;
			ParseItem lastParseItem = null;
			
			foreach(ParseItem pi in pis) 
			{
				if (parenthesesClosed) break;
			
				if (pi is Operand)
				{
					stack.Push(pi);
				}
				else if (pi is StringFunction)
				{
					handleFunction((StringFunction) pi, pis, stack);
				}
				else if (pi is Operator)
				{
					Operator op = (Operator) pi;
			
					// See if the operator is a binary or unary operator
					// It is a unary operator either if the stack is empty, or if
					// the last thing off the stack was another operator
					if (op is StringOperator)
					{
						StringOperator sop = (StringOperator) op;
						if (stack.Count<=0 || lastParseItem is Operator)
						{
							op = sop.UnaryOperator;
						}
						else
						{
							op = sop.BinaryOperator;
						}
					}
			
					if (operators.Count<=0)
					{
						// nothing much going on, so do nothing for the time being
						operators.Push(op);
					}
					else
					{
						Operator oper = (Operator) operators.Peek();
			
						// If the latest operator has a higher precedence then add to the 
						// operator stack and wait
						if (op.Precedence < oper.Precedence)
						{
							operators.Push(op);
						}
						else
						{
							// The operator is a lower precedence so we can sort out
							// some of the items on the stack
							operators.Pop(); // remove the operator from the stack
							oper.getOperands(stack);
							stack.Push(oper);
							operators.Push(op);
						}
					}
				}
				else if (pi is ArgumentSeparator)
				{
					// Clean up any remaining items on this stack
					while (operators.Count>0)
					{
						Operator o = (Operator) operators.Pop();
						o.getOperands(stack);
						stack.Push(o);
					}
			
					// Add it to the argument stack.  Create the argument stack
					// if necessary.  Items will be stored on the argument stack in
					// reverse order
					if (args == null)
					{
						args = new Stack();
					}
			
					args.Push(stack.Pop());
					// [TODO] check it
					//stack.empty();
					stack.Clear();
				}
				else if (pi is OpenParentheses)
				{
					ParseItem pi2 = parseCurrent(pis);
					Parenthesis p = new Parenthesis();
					pi2.Parent = p;
					p.add(pi2);
					stack.Push(p);
				}
				else if (pi is CloseParentheses)
				{
					parenthesesClosed = true;
				}
			
				lastParseItem = pi;
			}
			
			while (operators.Count > 0)
			{
				Operator o = (Operator) operators.Pop();
				o.getOperands(stack);
				stack.Push(o);
			}
			
			ParseItem rt = stack.Count > 0?(ParseItem) stack.Pop():null;
			
			// if the agument stack is not null, then add it to that stack
			// as well for good measure
			if (args != null && rt != null)
			{
				args.Push(rt);
			}
			
			arguments = args;
			
			if ((stack.Count > 0) || (operators.Count > 0))
			{
				logger.warn("Formula " + formula + " has a non-empty parse stack");
			}
			
			return rt;
		}
		
		/// <summary> Handles the case when parsing a string when a token is a function
		/// 
		/// </summary>
		/// <param name="sf">the string function
		/// </param>
		/// <param name="i"> the token iterator
		/// </param>
		/// <param name="stack">the parse tree stack
		/// </param>
		/// <exception cref=""> FormulaException if an error occurs
		/// </exception>
		private void  handleFunction(StringFunction sf, ArrayList pis, Stack stack)
		{
			ParseItem pi2 = parseCurrent(pis);
			
			// If the function is unknown, then throw an error
			if (sf.getFunction(settings) == Function.UNKNOWN)
			{
				throw new FormulaException(FormulaException.unrecognizedFunction);
			}
			
			// First check for possible optimized functions and possible
			// use of the Attribute token
			if (sf.getFunction(settings) == Function.SUM && arguments == null)
			{
				// this is handled by an attribute
				Attribute a = new Attribute(sf, settings);
				a.add(pi2);
				stack.Push(a);
				return ;
			}
			
			if (sf.getFunction(settings) == Function.IF)
			{
				// this is handled by an attribute
				Attribute a = new Attribute(sf, settings);
				
				// Add in the if conditions as a var arg function in
				// the correct order
				VariableArgFunction vaf = new VariableArgFunction(settings);
				
				// [TODO] TEST is the order the same as in Java?
				object[] argumentsArray = arguments.ToArray();
				for (int j = 0 ; j < argumentsArray.Length ; j++)
				{
				vaf.add((ParseItem) argumentsArray[j]);
				}
				
				
				a.IfConditions = vaf;
				stack.Push(a);
				return ;
			}
			
			// Function cannot be optimized.  See if it is a variable argument 
			// function or not
			if (sf.getFunction(settings).NumArgs == 0xff)
			{
				// If the arg stack has not been initialized, it means
				// that there was only one argument, which is the
				// returned parse item
				if (arguments == null)
				{
					VariableArgFunction vaf = new VariableArgFunction(sf.getFunction(settings), 1, settings);
					vaf.add(pi2);
					stack.Push(vaf);
				}
				else
				{
					// Add the args to the function in reverse order.  The 
					// VariableArgFunction will reverse these when it writes out the 
					// byte version as they are stored in the correct order
					// within Excel
					int numargs = arguments.Count;
					VariableArgFunction vaf = new VariableArgFunction(sf.getFunction(settings), numargs, settings);
					
					for (int j = 0; j < numargs; j++)
					{
						ParseItem pi3 = (ParseItem) arguments.Pop();
						vaf.add(pi3);
					}
					stack.Push(vaf);
					// [TODO] - check it
					//		arguments.empty();
					arguments.Clear();
					arguments = null;
				}
				return ;
			}
			
			// Function is a standard built in function
			BuiltInFunction bif = new BuiltInFunction(sf.getFunction(settings), settings);
			
			int numargs2 = sf.getFunction(settings).NumArgs;
			if (numargs2 == 1)
			{
				// only one item which is the returned ParseItem
				bif.add(pi2);
			}
			else
			{
				if ((arguments == null && numargs2 != 0) || (arguments != null && numargs2 != arguments.Count))
				{
					throw new FormulaException(FormulaException.incorrectArguments);
				}
				// multiple arguments so go to the arguments stack.  
				// Unlike the variable argument function, the args are
				// stored in reverse order
				// [TODO] TEST is the order the same as in Java?
				object[] argumentsArray = arguments.ToArray();
				for (int j = 0 ; j < numargs2 ; j++)
				{
				bif.add((ParseItem) argumentsArray[j]);
				}
				
			}
			stack.Push(bif);
		}
		
		/// <summary> Default behaviour is to do nothing
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
		static StringFormulaParser()
		{
			logger = Logger.getLogger(typeof(StringFormulaParser));
		}
	}
}
