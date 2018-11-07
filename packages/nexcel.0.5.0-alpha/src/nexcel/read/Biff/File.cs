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
using NExcel;
using NExcel.Biff;
using NExcelUtils;

namespace NExcel.Read.Biff
{
	
	/// <summary> File containing the data from the binary stream</summary>
	public class File 
	{
		/// <summary> Gets the position in the stream
		/// 
		/// </summary>
		/// <returns> the position in the stream
		/// </returns>
		/// <summary> Saves the current position and temporarily sets the position to be the
		/// new one.  The original position may be restored usind the restorePos()
		/// method. This is used when reading in the cell values of the sheet - an
		/// addition in 1.6 for memory allocation reasons.
		/// 
		/// These methods are used by the SheetImpl.readSheet() when it is reading
		/// in all the cell values
		/// 
		/// </summary>
		/// <param name="p">the temporary position
		/// </param>
		virtual public int Pos
		{
			get
			{
				return filePos;
			}
			
			set
			{
				oldPos = filePos;
				filePos = value;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The data from the excel 97 file</summary>
		private sbyte[] data;
		/// <summary> The current position within the file</summary>
		private int filePos;
		/// <summary> The saved pos</summary>
		private int oldPos;
		/// <summary> The initial file size</summary>
		private int initialFileSize;
		/// <summary> The amount to increase the growable array by</summary>
		private int arrayGrowSize;
		/// <summary> The workbook settings</summary>
		private WorkbookSettings workbookSettings;
		
		/// <summary> Constructs a file from the input stream
		/// 
		/// </summary>
		/// <param name="is">the input stream
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <exception cref=""> IOException
		/// </exception>
		/// <exception cref=""> BiffException
		/// </exception>
		public File(System.IO.Stream stream, WorkbookSettings ws)
		{
			// Initialize the file sizing parameters from the settings
			workbookSettings = ws;
			initialFileSize = workbookSettings.InitialFileSize;
			arrayGrowSize = workbookSettings.ArrayGrowSize;
			
			sbyte[] d = new sbyte[initialFileSize];
			int bytesRead = NExcelUtils.File.ReadInput(stream, ref d, 0, d.Length);
			int pos = bytesRead;
			
			// Handle thread interruptions, in case the user keeps pressing
			// the Submit button from a browser.  Thanks to Mike Smith for this
			// [TODO-NExcel_Next] - make the same in C#
			//    if (Thread.currentThread().isInterrupted())
			//    {
			//      throw new InterruptedIOException();
			//    }
			
			while (bytesRead != - 1)
			{
				if (pos >= d.Length)
				{
					// Grow the array
					sbyte[] newArray = new sbyte[d.Length + arrayGrowSize];
					Array.Copy(d, 0, newArray, 0, d.Length);
					d = newArray;
				}
				bytesRead = NExcelUtils.File.ReadInput(stream, ref d, pos, d.Length - pos);
				pos += bytesRead;
				
				// [TODO-NExcel_Next] - make the same in C#
				//	  if (Thread.currentThread().isInterrupted())
				//      {
				//        throw new InterruptedIOException();
				//      }
			}
			
			bytesRead = pos + 1;
			
			// Perform file reading checks and throw exceptions as necessary
			if (bytesRead == 0)
			{
				throw new BiffException(BiffException.excelFileNotFound);
			}
			
			CompoundFile cf = new CompoundFile(d, ws);
			try
			{
				data = cf.getStream("workbook");
			}
			catch (BiffException e)
			{
				// this might be in excel 95 format - try again
				data = cf.getStream("book");
			}
			cf = null;
			
			if (!workbookSettings.GCDisabled)
			{
				System.GC.Collect();
			}
			
			// Uncomment the following lines to send the pure workbook stream
			// (ie. a defragged ole stream) to an output file
			
			//       FileOutputStream fos = new FileOutputStream("defraggedxls");
			//       fos.write(data);
			//       fos.close();
		}
		
		/// <summary> Returns the next data record and increments the pointer
		/// 
		/// </summary>
		/// <returns> the next data record
		/// </returns>
		internal virtual Record next()
		{
			Record r = new Record(data, filePos, this);
			return r;
		}
		
		/// <summary> Peek ahead to the next record, without incrementing the file position
		/// 
		/// </summary>
		/// <returns> the next record
		/// </returns>
		internal virtual Record peek()
		{
			int tempPos = filePos;
			Record r = new Record(data, filePos, this);
			filePos = tempPos;
			return r;
		}
		
		/// <summary> Skips forward the specified number of bytes
		/// 
		/// </summary>
		/// <param name="bytes">the number of bytes to skip forward
		/// </param>
		public virtual void  skip(int bytes)
		{
			filePos += bytes;
		}
		
		/// <summary> Copies the bytes into a new array and returns it.
		/// 
		/// </summary>
		/// <param name="pos">the position to read from
		/// </param>
		/// <param name=".Length">the number of bytes to read
		/// </param>
		/// <returns> The bytes read
		/// </returns>
		public virtual sbyte[] read(int pos, int length)
		{
			sbyte[] ret = new sbyte[length];
			Array.Copy(data, pos, ret, 0, length);
			return ret;
		}
		
		/// <summary> Restores the original position
		/// 
		/// These methods are used by the SheetImpl.readSheet() when it is reading
		/// in all the cell values
		/// </summary>
		public virtual void  restorePos()
		{
			filePos = oldPos;
		}
		
		/// <summary> Moves to the first bof in the file</summary>
		private void  moveToFirstBof()
		{
			bool bofFound = false;
			while (!bofFound)
			{
				int code = IntegerHelper.getInt(data[filePos], data[filePos + 1]);
				if (code == NExcel.Biff.Type.BOF.Value)
				{
					bofFound = true;
				}
				else
				{
					skip(128);
				}
			}
		}
		
		/// <summary> "Closes" the biff file
		/// 
		/// </summary>
		/// <deprecated> As of version 1.6 use workbook.close() instead
		/// </deprecated>
		public virtual void  close()
		{
		}
		
		/// <summary> Clears the contents of the file</summary>
		public virtual void  clear()
		{
			data = null;
		}
		
		/// <summary> Determines if the current position exceeds the end of the file
		/// 
		/// </summary>
		/// <returns> TRUE if there is more data left in the array, FALSE otherwise
		/// </returns>
		public virtual bool hasNext()
		{
			// Allow four bytes for the record code and its .Length
			return filePos < data.Length - 4;
		}
		static File()
		{
			logger = Logger.getLogger(typeof(File));
		}
	}
}
