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
namespace NExcel.Biff.Drawing
{
	
	class ClientAnchor:EscherAtom
	{
		virtual internal double X1
		{
			get
			{
				return x1;
			}
			
		}
		virtual internal double Y1
		{
			get
			{
				return y1;
			}
			
		}
		virtual internal double X2
		{
			get
			{
				return x2;
			}
			
		}
		virtual internal double Y2
		{
			get
			{
				return y2;
			}
			
		}
		new private sbyte[] data;
		
		private double x1;
		private double y1;
		private double x2;
		private double y2;
		
		public ClientAnchor(EscherRecordData erd):base(erd)
		{
			sbyte[] bytes = Bytes;
			
			// The x1 cell
			int x1Cell = IntegerHelper.getInt(bytes[2], bytes[3]);
			int x1Fraction = IntegerHelper.getInt(bytes[4], bytes[5]);
			
			x1 = x1Cell + (double) x1Fraction / (double) 1024;
			
			// The y1 cell
			int y1Cell = IntegerHelper.getInt(bytes[6], bytes[7]);
			int y1Fraction = IntegerHelper.getInt(bytes[8], bytes[9]);
			
			y1 = y1Cell + (double) y1Fraction / (double) 256;
			
			// The x2 cell
			int x2Cell = IntegerHelper.getInt(bytes[10], bytes[11]);
			int x2Fraction = IntegerHelper.getInt(bytes[12], bytes[13]);
			
			x2 = x2Cell + (double) x2Fraction / (double) 1024;
			
			// The y1 cell
			int y2Cell = IntegerHelper.getInt(bytes[14], bytes[15]);
			int y2Fraction = IntegerHelper.getInt(bytes[16], bytes[17]);
			
			y2 = y2Cell + (double) y2Fraction / (double) 256;
		}
		
		public ClientAnchor(double x1, double y1, double x2, double y2):base(EscherRecordType.CLIENT_ANCHOR)
		{
			this.x1 = x1;
			this.y1 = y1;
			this.x2 = x2;
			this.y2 = y2;
		}
		
		public override sbyte[] Data
		{
			get
			{
				data = new sbyte[18];
				IntegerHelper.getTwoBytes(0x2, data, 0);
			
				// The x1 cell
				IntegerHelper.getTwoBytes((int) x1, data, 2);
			
				// The x1 fraction into the cell 0-1024
				int x1fraction = (int) ((x1 - (int) x1) * 1024);
				IntegerHelper.getTwoBytes(x1fraction, data, 4);
			
				// The y1 cell
				IntegerHelper.getTwoBytes((int) y1, data, 6);
			
				// The y1 fraction into the cell 0-256
				int y1fraction = (int) ((y1 - (int) y1) * 256);
				IntegerHelper.getTwoBytes(y1fraction, data, 8);
			
				// The x2 cell
				IntegerHelper.getTwoBytes((int) x2, data, 10);
			
				// The x2 fraction into the cell 0-1024
				int x2fraction = (int) ((x2 - (int) x2) * 1024);
				IntegerHelper.getTwoBytes(x2fraction, data, 12);
			
				// The y2 cell
				IntegerHelper.getTwoBytes((int) y2, data, 14);
			
				// The y2 fraction into the cell 0-256
				int y2fraction = (int) ((y2 - (int) y2) * 256);
				IntegerHelper.getTwoBytes(y2fraction, data, 16);
			
				return setHeaderData(data);
			}
		}
	}
}
