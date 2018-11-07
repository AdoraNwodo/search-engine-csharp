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
using NExcel.Biff;
namespace NExcel.Read.Biff
{
	
	/// <summary> Contains the page set up for a sheet</summary>
	public class SetupRecord:RecordData
	{
		/// <summary> Accessor for the orientation.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> TRUE if the orientation is portrait, FALSE if it is landscape
		/// </returns>
		virtual public bool isPortrait()
		{
				return portraitOrientation;
		}


		/// <summary> Accessor for the header.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> the header margin
		/// </returns>
		virtual public double HeaderMargin
		{
			get
			{
				return headerMargin;
			}
			
		}
		/// <summary> Accessor for the footer.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> the footer margin
		/// </returns>
		virtual public double FooterMargin
		{
			get
			{
				return footerMargin;
			}
			
		}
		/// <summary> Accessor for the paper size.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> the footer margin
		/// </returns>
		virtual public int PaperSize
		{
			get
			{
				return paperSize;
			}
			
		}
		/// <summary> Accessor for the scale factor.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> the scale factor
		/// </returns>
		virtual public int ScaleFactor
		{
			get
			{
				return scaleFactor;
			}
			
		}
		/// <summary> Accessor for the page height.  called when copying sheets
		/// 
		/// </summary>
		/// <returns> the page to start printing at
		/// </returns>
		virtual public int PageStart
		{
			get
			{
				return pageStart;
			}
			
		}
		/// <summary> Accessor for the fit width.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> the fit width
		/// </returns>
		virtual public int FitWidth
		{
			get
			{
				return fitWidth;
			}
			
		}
		/// <summary> Accessor for the fit height.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> the fit height
		/// </returns>
		virtual public int FitHeight
		{
			get
			{
				return fitHeight;
			}
			
		}
		/// <summary> The horizontal print resolution.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> the horizontal print resolution
		/// </returns>
		virtual public int HorizontalPrintResolution
		{
			get
			{
				return horizontalPrintResolution;
			}
			
		}
		/// <summary> Accessor for the vertical print resolution.  Called when copying sheets
		/// 
		/// </summary>
		/// <returns> an vertical print resolution
		/// </returns>
		virtual public int VerticalPrintResolution
		{
			get
			{
				return verticalPrintResolution;
			}
			
		}
		/// <summary> Accessor for the number of copies
		/// 
		/// </summary>
		/// <returns> the number of copies
		/// </returns>
		virtual public int Copies
		{
			get
			{
				return copies;
			}
			
		}
		/// <summary> The raw data</summary>
		private sbyte[] data;
		
		/// <summary> The orientation flag</summary>
		private bool portraitOrientation;
		
		/// <summary> The header margin</summary>
		private double headerMargin;
		
		/// <summary> The footer margin</summary>
		private double footerMargin;
		
		/// <summary> The paper size</summary>
		private int paperSize;
		
		/// <summary> The scale factor</summary>
		private int scaleFactor;
		
		/// <summary> The page start</summary>
		private int pageStart;
		
		/// <summary> The fit width</summary>
		private int fitWidth;
		
		/// <summary> The fit height</summary>
		private int fitHeight;
		
		/// <summary> The horizontal print resolution</summary>
		private int horizontalPrintResolution;
		
		/// <summary> The vertical print resolution</summary>
		private int verticalPrintResolution;
		
		/// <summary> The number of copies</summary>
		private int copies;
		
		/// <summary> Constructor which creates this object from the binary data
		/// 
		/// </summary>
		/// <param name="t">the record
		/// </param>
		internal SetupRecord(Record t):base(NExcel.Biff.Type.SETUP)
		{
			
			data = t.Data;
			
			paperSize = IntegerHelper.getInt(data[0], data[1]);
			scaleFactor = IntegerHelper.getInt(data[2], data[3]);
			pageStart = IntegerHelper.getInt(data[4], data[5]);
			fitWidth = IntegerHelper.getInt(data[6], data[7]);
			fitHeight = IntegerHelper.getInt(data[8], data[9]);
			horizontalPrintResolution = IntegerHelper.getInt(data[12], data[13]);
			verticalPrintResolution = IntegerHelper.getInt(data[14], data[15]);
			copies = IntegerHelper.getInt(data[32], data[33]);
			
			headerMargin = DoubleHelper.getIEEEDouble(data, 16);
			footerMargin = DoubleHelper.getIEEEDouble(data, 24);
			
			
			
			int grbit = IntegerHelper.getInt(data[10], data[11]);
			portraitOrientation = ((grbit & 0x02) != 0);
		}
	}
}
