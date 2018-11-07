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
using NExcel.Read.Biff;
namespace NExcel.Biff.Drawing
{
	
	/// <summary> Contains the various biff records used to insert a chart into a 
	/// worksheet
	/// </summary>
	public class Chart : ByteData, EscherStream
	{
		/// <summary> Gets the SpContainer containing the charts drawing information
		/// 
		/// </summary>
		/// <returns> the spContainer
		/// </returns>
		virtual internal SpContainer SpContainer
		{
			get
			{
				EscherRecordData er = new EscherRecordData(this, 0);
				Assert.verify(er.isContainer());
				
				EscherContainer escherData = new EscherContainer(er);
				
				SpContainer spContainer = null;
				if (escherData.Type == EscherRecordType.DG_CONTAINER)
				{
					EscherRecordData erd = new EscherRecordData(this, 80);
					Assert.verify(erd.Type == EscherRecordType.SP_CONTAINER);
					spContainer = new SpContainer(erd);
				}
				else
				{
					Assert.verify(escherData.Type == EscherRecordType.SP_CONTAINER);
					spContainer = new SpContainer(er);
				}
				
				return spContainer;
			}
			
		}
		/// <summary> Accessor for the mso drawing record
		/// 
		/// </summary>
		/// <returns> the drawing record
		/// </returns>
		virtual internal MsoDrawingRecord MsoDrawingRecord
		{
			get
			{
				return msoDrawingRecord;
			}
			
		}
		/// <summary> Accessor for the obj record
		/// 
		/// </summary>
		/// <returns> the obj record
		/// </returns>
		virtual internal ObjRecord ObjRecord
		{
			get
			{
				return objRecord;
			}
			
		}
		/// <summary> The MsoDrawingRecord associated with the chart</summary>
		private MsoDrawingRecord msoDrawingRecord;
		
		/// <summary> The ObjRecord associated with the chart</summary>
		private ObjRecord objRecord;
		
		/// <summary> The start pos of the chart bof stream in the data file</summary>
		private int startpos;
		
		/// <summary> The start pos of the chart bof stream in the data file</summary>
		private int endpos;
		
		/// <summary> A handle to the Excel file</summary>
		private File file;
		
		/// <summary> The chart byte data</summary>
		private sbyte[] data;
		
		/// <summary> Flag which indicates that the byte data has been initialized</summary>
		private bool initialized;
		
		/// <summary> The workbook settings</summary>
		private WorkbookSettings workbookSettings;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="mso">a <code>MsoDrawingRecord</code> value
		/// </param>
		/// <param name="obj">an <code>ObjRecord</code> value
		/// </param>
		/// <param name="sp">an <code>int</code> value
		/// </param>
		/// <param name="ep">an <code>int</code> value
		/// </param>
		/// <param name="f">a <code>File</code> value
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public Chart(MsoDrawingRecord mso, ObjRecord obj, int sp, int ep, File f, WorkbookSettings ws)
		{
			msoDrawingRecord = mso;
			objRecord = obj;
			startpos = sp;
			endpos = ep;
			file = f;
			workbookSettings = ws;
			initialized = false;
			
			// Note:  mso and obj values can be null if we are creating a chart
			// which takes up an entire worksheet.  Check that both are null or both
			// not null though
			Assert.verify((mso != null && obj != null) || (mso == null && obj == null));
		}
		
		/// <summary> Gets the entire binary record for the chart as a chunk of binary data
		/// 
		/// </summary>
		/// <returns> the bytes
		/// </returns>
		public virtual sbyte[] getBytes()
		{
			if (!initialized)
			{
				initialize();
			}
			
			return data;
		}
		
		/// <summary> Implementation of the EscherStream method
		/// 
		/// </summary>
		/// <returns> the data
		/// </returns>
		public virtual sbyte[] getData()
		{
			return msoDrawingRecord.getRecord().Data;
		}
		
		/// <summary> Initializes the charts byte data</summary>
		private void  initialize()
		{
			data = file.read(startpos, endpos - startpos);
			initialized = true;
		}
		
		/// <summary> Rationalizes the sheets xf index mapping</summary>
		/// <param name="xfMapping">the index mapping for XFRecords
		/// </param>
		/// <param name="fontMapping">the index mapping for fonts
		/// </param>
		/// <param name="formatMapping">the index mapping for formats
		/// </param>
		public virtual void  rationalize(IndexMapping xfMapping, IndexMapping fontMapping, IndexMapping formatMapping)
		{
			if (!initialized)
			{
				initialize();
			}
			
			// Read through the array, looking for the data types
			// This is a total hack bodge for now - it will eventually need to be
			// integrated properly
			int pos = 0;
			int code = 0;
			int length = 0;
			NExcel.Biff.Type type = null;
			while (pos < data.Length)
			{
				code = IntegerHelper.getInt(data[pos], data[pos + 1]);
				length = IntegerHelper.getInt(data[pos + 2], data[pos + 3]);
				
				type = NExcel.Biff.Type.getType(code);
				
				if (type == NExcel.Biff.Type.FONTX)
				{
					int fontind = IntegerHelper.getInt(data[pos + 4], data[pos + 5]);
					IntegerHelper.getTwoBytes(fontMapping.getNewIndex(fontind), data, pos + 4);
				}
				else if (type == NExcel.Biff.Type.FBI)
				{
					int fontind = IntegerHelper.getInt(data[pos + 12], data[pos + 13]);
					IntegerHelper.getTwoBytes(fontMapping.getNewIndex(fontind), data, pos + 12);
				}
				else if (type == NExcel.Biff.Type.IFMT)
				{
					int formind = IntegerHelper.getInt(data[pos + 4], data[pos + 5]);
					IntegerHelper.getTwoBytes(formatMapping.getNewIndex(formind), data, pos + 4);
				}
				
				pos += length + 4;
			}
		}
	}
}
