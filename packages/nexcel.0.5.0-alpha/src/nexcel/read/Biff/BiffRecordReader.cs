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
namespace NExcel.Read.Biff
{
	
	/// <summary> Serves up Record objects from a biff file.  This object is used by the
	/// demo programs BiffDump and ... only and has no influence whatsoever on
	/// the JExcelApi reading and writing of excel sheets
	/// </summary>
	public class BiffRecordReader
	{
		/// <summary> Gets the position of the current record in the biff file
		/// 
		/// </summary>
		/// <returns> the position
		/// </returns>
		virtual public int Pos
		{
			get
			{
				return file.Pos - record.Length - 4;
			}
			
		}
		/// <summary> The biff file</summary>
		private File file;
		
		/// <summary> The current record retrieved</summary>
		private Record record;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="f">the biff file
		/// </param>
		public BiffRecordReader(File f)
		{
			file = f;
		}
		
		/// <summary> Sees if there are any more records to read
		/// 
		/// </summary>
		/// <returns> TRUE if there are more records, FALSE otherwise
		/// </returns>
		public virtual bool hasNext()
		{
			return file.hasNext();
		}
		
		/// <summary> Gets the next record
		/// 
		/// </summary>
		/// <returns> the next record
		/// </returns>
		public virtual Record next()
		{
			record = file.next();
			return record;
		}
	}
}
