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
	
	/// <summary> Exception thrown when parsing a formula</summary>
	public class FormulaException:JXLException
	{
		public class FormulaMessage
		{
			/// <summary> The message</summary>
			public string message;
			/// <summary> Constructs this exception with the specified message
			/// 
			/// </summary>
			/// <param name="m">the message
			/// </param>
			internal FormulaMessage(string m)
			{
				message = m;
			}
		}
		
		
		internal static FormulaMessage unrecognizedToken;
		
		
		internal static FormulaMessage unrecognizedFunction;
		
		
		internal static FormulaMessage biff8Supported;
		
		internal static FormulaMessage lexicalError;
		
		internal static FormulaMessage incorrectArguments;
		
		internal static FormulaMessage sheetRefNotFound;
		
		internal static FormulaMessage cellNameNotFound;
		
		
		/// <summary> Constructs this exception with the specified message
		/// 
		/// </summary>
		/// <param name="m">the message
		/// </param>
		public FormulaException(FormulaMessage m):base(m.message)
		{
		}
		
		/// <summary> Constructs this exception with the specified message
		/// 
		/// </summary>
		/// <param name="m">the message
		/// </param>
		public FormulaException(FormulaMessage m, int val):base(m.message + " " + val)
		{
		}
		
		/// <summary> Constructs this exception with the specified message
		/// 
		/// </summary>
		/// <param name="m">the message
		/// </param>
		public FormulaException(FormulaMessage m, string val):base(m.message + " " + val)
		{
		}
		static FormulaException()
		{
			unrecognizedToken = new FormulaMessage("Unrecognized token");
			unrecognizedFunction = new FormulaMessage("Unrecognized function");
			biff8Supported = new FormulaMessage("Only biff8 formulas are supported");
			lexicalError = new FormulaMessage("Lexical error:  ");
			incorrectArguments = new FormulaMessage("Incorrect arguments supplied to function");
			sheetRefNotFound = new FormulaMessage("Could not find sheet");
			cellNameNotFound = new FormulaMessage("Could not find named cell");
		}
	}
}
