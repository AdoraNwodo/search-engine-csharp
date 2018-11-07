/// <summary>******************************************************************
/// 
/// Copyright (C) 2005  Stefano Franco
///
/// Based on JExcelAPI by Andrew Khan.
/// 
/// This library is free software; you can redistribute it and/or
/// modify it under the terms of the GNU Library General Public
/// License as published by the Free Software Foundation; either
/// version 2 of the License, or (at your option) any later version.
/// 
/// This library is distributed in the hope that it will be useful,
/// but WITHOUT ANY WARRANTY; without even the implied warranty of
/// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
/// Library General Public License for more details.
/// 
/// You should have received a copy of the GNU Library General Public
/// License along with this library; if not, write to the Free Software
/// Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
/// *************************************************************************
/// </summary>
using System;
namespace common
{
	
	/// <summary> An exception thrown when an assert (from the Assert class) fails</summary>
	public class AssertionFailed:System.SystemException
	{
		/// <summary> Default constructor
		/// Prints the stack trace
		/// </summary>
		public AssertionFailed():base()
		{
			//    printStackTrace();
		}
		
		/// <summary> Constructor with message
		/// Prints the stack trace
		/// 
		/// </summary>
		/// <param name="s">Message thrown with the assertion
		/// </param>
		public AssertionFailed(string s):base(s)
		{
		}
	}
}