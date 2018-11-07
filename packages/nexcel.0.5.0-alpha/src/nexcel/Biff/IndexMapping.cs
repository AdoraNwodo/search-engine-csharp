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
using common;
namespace NExcel.Biff
{
	
	/// <summary> This class is a wrapper for a list of mappings between indices.
	/// It is used when removing duplicate records and specifies the new
	/// index for cells which have the duplicate format
	/// </summary>
	public sealed class IndexMapping
	{
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The array of new indexes for an old one</summary>
		private int[] newIndices;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="size">the number of index numbers to be mapped
		/// </param>
		internal IndexMapping(int size)
		{
			newIndices = new int[size];
		}
		
		/// <summary> Sets a mapping</summary>
		/// <param name="oldIndex">the old index
		/// </param>
		/// <param name="newIndex">the new index
		/// </param>
		internal void  setMapping(int oldIndex, int newIndex)
		{
			newIndices[oldIndex] = newIndex;
		}
		
		/// <summary> Gets the new cell format index</summary>
		/// <param name="oldIndex">the existing index number
		/// </param>
		/// <returns> the new index number
		/// </returns>
		public int getNewIndex(int oldIndex)
		{
			return newIndices[oldIndex];
		}
		static IndexMapping()
		{
			logger = Logger.getLogger(typeof(IndexMapping));
		}
	}
}
