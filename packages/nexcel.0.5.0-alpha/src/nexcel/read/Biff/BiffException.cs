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
namespace NExcel.Read.Biff
{
	
	/// <summary> Exception thrown when reading a biff file</summary>
	public class BiffException:JXLException
	{
		/// <summary> Inner class containing the various error messages</summary>
		public class BiffMessage
		{
			/// <summary> The formatted message</summary>
			public string message;
			/// <summary> Constructs this exception with the specified message
			/// 
			/// </summary>
			/// <param name="m">the messageA
			/// </param>
			internal BiffMessage(string m)
			{
				message = m;
			}
		}
		
		
		internal static readonly BiffMessage unrecognizedBiffVersion = new BiffMessage("Unrecognized biff version");
		
		
		internal static readonly BiffMessage expectedGlobals = new BiffMessage("Expected globals");
		
		
		internal static readonly BiffMessage excelFileTooBig = new BiffMessage("Warning:  not all of the excel file could be read");
		
		
		internal static readonly BiffMessage excelFileNotFound = new BiffMessage("The input file was not found");
		
		
		internal static readonly BiffMessage unrecognizedOLEFile = new BiffMessage("Unable to recognize OLE stream");
		
		
		internal static readonly BiffMessage streamNotFound = new BiffMessage("Compound file does not contain the specified stream");
		
		
		internal static readonly BiffMessage passwordProtected = new BiffMessage("The workbook is password protected");
		
		/// <summary> Constructs this exception with the specified message
		/// 
		/// </summary>
		/// <param name="m">the message
		/// </param>
		public BiffException(BiffMessage m):base(m.message)
		{
		}
	}
}
