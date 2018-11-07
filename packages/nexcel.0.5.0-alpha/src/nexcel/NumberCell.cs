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
using NumberFormatInfo = NExcelUtils.NumberFormatInfo;
namespace NExcel
{
	
	
	/// <summary> A cell which contains a numerical value</summary>
	public interface NumberCell : Cell
		{
			/// <summary> Gets the double value for this cell.
			/// 
			/// </summary>
			/// <returns> the cell value
			/// </returns>
			double DoubleValue
			{
				get;
				
			}
			/// <summary> Gets the NumberFormat used to format this cell.  This is the java
			/// equivalent of the Excel format
			/// 
			/// </summary>
			/// <returns> the NumberFormat used to format the cell
			/// </returns>
			NumberFormatInfo NumberFormat
			{
				//  public NumberFormat getNumberFormat();
				
				get;
				
			}
		}
}
