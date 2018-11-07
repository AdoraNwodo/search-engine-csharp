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
//using System.Resources;
using common;
using NExcelUtils;


namespace NExcel.Biff.Formula
{
	
	/// <summary> A class which contains the function names for the current workbook. The
	/// function names can potentially vary from workbook to workbook depending
	/// on the locale
	/// </summary>
	public class FunctionNames
	{
		/// <summary> The logger class</summary>
		private static Logger logger;
		
		/// <summary> A hash mapping keyed on the function and returning its locale specific 
		/// name
		/// </summary>
		private Hashtable names;
		
		/// <summary> A hash mapping keyed on the locale specific name and returning the 
		/// function
		/// </summary>
		private Hashtable functions;
		
		/// <summary> Constructor
		/// @ws the workbook settings
		/// </summary>
		public FunctionNames(System.Globalization.CultureInfo l)
		{
			
			ResourceManager rm = new ResourceManager("NExcel.Biff.Formula.FunctionNames", l, this.GetType().Assembly);

			names = new Hashtable(Function.functions.Length);
			functions = new Hashtable(Function.functions.Length);
			
			// Iterate through all the functions, adding them to the hash maps
			Function f = null;
			string n = null;
			string propname = null;
			for (int i = 0; i < Function.functions.Length; i++)
			{
				f = Function.functions[i];
				propname = f.PropertyName;
				
				n = propname.Length != 0 ? rm.GetString(propname) : null;
				
				
				if ((System.Object) n != null)
				{
					names[f] =  n;
					functions[n] =  f;
				}
			}
		}
		
		/// <summary> Gets the function for the specified name</summary>
		internal virtual Function getFunction(string s)
		{
			return (Function) functions[s];
		}
		
		/// <summary> Gets the name for the function</summary>
		internal virtual string getName(Function f)
		{
			return (string) names[f];
		}
		static FunctionNames()
		{
			logger = Logger.getLogger(typeof(FunctionNames));
		}
	}
}
