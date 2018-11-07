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
namespace NExcel.Biff
{
	
	/// <summary> The interface implemented by the various number and date format styles.
	/// The methods on this interface are called internally when generating a
	/// spreadsheet
	/// </summary>
	public interface DisplayFormat
		{
			/// <summary> Accessor for the index style of this format
			/// 
			/// </summary>
			/// <returns> the index for this format
			/// </returns>
			int FormatIndex
			{
				get;
				
			}
			/// <summary> Accessor to see whether this format has been initialized
			/// 
			/// </summary>
			/// <returns> TRUE if initialized, FALSE otherwise
			/// </returns>
			bool isInitialized();

			/// <summary> Accessor to determine whether or not this format is built in
			/// 
			/// </summary>
			/// <returns> TRUE if this format is a built in format, FALSE otherwise
			/// </returns>
			bool isBuiltIn();
			
			/// <summary> Initializes this format with the specified index number
			/// 
			/// </summary>
			/// <param name="pos">the position of this format record in the workbook
			/// </param>
			void  initialize(int pos);
		}
}
