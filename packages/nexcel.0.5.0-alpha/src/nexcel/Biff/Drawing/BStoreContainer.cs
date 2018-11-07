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
	
	
	/// <summary> A BStoreContainer escher record</summary>
	class BStoreContainer:EscherContainer
	{
		/// <summary> Accessor for the number of blips
		/// 
		/// </summary>
		/// <returns> the number of blips
		/// </returns>
		virtual public int NumBlips
		{
			get
			{
				return numBlips;
			}
			
			/// <summary> Accessor for the drawing
			/// 
			/// </summary>
			/// <param name="i">the index number of the drawing to return
			/// </param>
			// [NON USATA]
			//  public void getDrawing(int i)
			//  {
			//    EscherRecord[] children = getChildren();
			//    BlipStoreEntry bse = (BlipStoreEntry) children[i];
			//  }
			
		}
		/// <summary> The number of blips inside this container</summary>
		private int numBlips;
		
		/// <summary> Constructor used to instantiate this object when reading from an
		/// escher stream
		/// 
		/// </summary>
		/// <param name="erd">the escher data
		/// </param>
		public BStoreContainer(EscherRecordData erd):base(erd)
		{
			numBlips = Instance;
		}
		
		/// <summary> Constructor used when writing out an escher record
		/// 
		/// </summary>
		/// <param name="count">the number of blips
		/// </param>
		public BStoreContainer(int count):base(EscherRecordType.BSTORE_CONTAINER)
		{
			numBlips = count;
			Instance = numBlips;
		}
	}
}
