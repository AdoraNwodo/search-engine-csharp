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
	
	/// <summary> Simple assertion mechanism for use during development</summary>
	public sealed class Assert
	{
		/// <summary> Throws an AssertionFailed exception if the specified condition is
		/// false
		/// 
		/// </summary>
		/// <param name="condition">The assertion condition which must be true
		/// </param>
		public static void  verify(bool condition)
		{
			if (!condition)
			{
				throw new AssertionFailed();
			}
		}
		
		/// <summary> If the condition evaluates to false, an AssertionFailed is thrown
		/// 
		/// </summary>
		/// <param name="message">A message thrown with the failed assertion
		/// </param>
		/// <param name="condition">If this evaluates to false, an AssertionFailed is thrown
		/// </param>
		public static void  verify(bool condition, string message)
		{
			if (!condition)
			{
				throw new AssertionFailed(message);
			}
		}
	}
}