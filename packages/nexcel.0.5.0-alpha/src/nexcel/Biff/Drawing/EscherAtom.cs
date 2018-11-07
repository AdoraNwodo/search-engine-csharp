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
namespace NExcel.Biff.Drawing
{
	
	/// <summary> Class for atoms.  This may be instantiated as is for unknown/uncared about
	/// atoms, or subclassed if we have some semantic interest in the contents
	/// </summary>
	class EscherAtom:EscherRecord
	{
		override public sbyte[] Data
		{
			get
			{
				//System.Console.Error.WriteLine("WARNING:  Escher atom getData called");
				sbyte[] data = new sbyte[0];
				return setHeaderData(data);
			}
			
		}
		public EscherAtom(EscherRecordData erd):base(erd)
		{
		}
		
		protected internal EscherAtom(EscherRecordType type):base(type)
		{
		}
	}
}
