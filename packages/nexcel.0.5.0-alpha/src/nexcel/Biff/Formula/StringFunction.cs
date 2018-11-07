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
	/// <summary> Class used to hold a function when reading it in from a string.  At this
	/// stage it is unknown whether it is a BuiltInFunction or a VariableArgFunction
	/// </summary>
	class StringFunction:StringParseItem
	{
		/// <summary> The function</summary>
		private Function function;
		
		/// <summary> The function string</summary>
		private string functionString;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="s">the lexically parsed stirng
		/// </param>
		internal StringFunction(string s)
		{
			functionString = s.Substring(0, (s.Length - 1) - (0));
		}
		
		/// <summary> Accessor for the function
		/// 
		/// </summary>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <returns> the function
		/// </returns>
		internal virtual Function getFunction(WorkbookSettings ws)
		{
			if (function == null)
			{
				function = Function.getFunction(functionString, ws);
			}
			return function;
		}
	}
}
