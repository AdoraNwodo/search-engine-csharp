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
using System.Collections;
using common;
using NExcel;
using NExcel.Biff;
namespace NExcel.Biff.Drawing
{
	// [TODO-NExcel_Next]
	//import NExcel.Write.biff.File;
	
	/// <summary> Handles the writing out of the different charts and images on a sheet.  
	/// Called by the SheetWriter object
	/// </summary>
	public class SheetDrawingWriter
	{
		/// <summary> The drawings on the sheet
		/// 
		/// </summary>
		/// <param name="dr">the list of drawings
		/// </param>
		virtual public ArrayList Drawings
		{
			set
			{
				drawings = value;
			}
			
		}
		/// <summary> Accessor for the charts on the sheet
		/// 
		/// </summary>
		/// <returns> the charts
		/// </returns>
		/// <summary> Sets the charts on the sheet
		/// 
		/// </summary>
		/// <param name="ch">the charts
		/// </param>
		virtual public Chart[] Charts
		{
			get
			{
				return charts;
			}
			
			set
			{
				charts = value;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The drawings on the sheet</summary>
		private ArrayList drawings;
		
		/// <summary> The charts on the sheet</summary>
		private Chart[] charts;
		
		/// <summary> The workbook settings</summary>
		private WorkbookSettings workbookSettings;
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="ws">the workbook settings
		/// </param>
		public SheetDrawingWriter(WorkbookSettings ws)
		{
			charts = new Chart[0];
		}
		
		// [TODO-NExcel_Next]
		//  /**
		//   * Writes out the MsoDrawing records and Obj records for each image
		//   * and chart on the sheet
		//   *
		//   * @param outputFile
		//   * @exception IOException
		//   */
		//  public void write(File outputFile) throws IOException
		//  {
		//    // If there are no drawings or charts on this sheet then exit
		//    if (drawings.size() == 0 && charts.Length == 0)
		//    {
		//      return;
		//    }
		//
		//    // See if any drawing has been modified
		//    boolean modified = false;
		//    for (Iterator i = drawings.iterator() ; i.hasNext() && !modified ;)
		//    {
		//      Drawing d = (Drawing) i.next();
		//      if (d.getOrigin() != Drawing.READ)
		//      {
		//        modified = true;
		//      }
		//    }
		//
		//    // If no drawing has been modified, then simply write them straight out 
		//    // again and exit
		//    if (!modified)
		//    {
		//      writeUnmodified(outputFile);
		//      return;
		//    }
		//
		//    int numImages = drawings.size();
		//    Object[] spContainerData = new Object[numImages + charts.Length];
		//    int .Length = 0;
		//    EscherContainer firstSpContainer = null;
		//
		//    // Get all the spContainer byte data from the drawings 
		//    // and store in an array
		//    for (int i = 0 ; i < numImages; i++)
		//    {
		//      Drawing drawing = (Drawing) drawings.get(i);
		//    
		//      SpContainer spc = drawing.getSpContainer();
		//      byte[] data = spc.getData();
		//      spContainerData[i] = data;
		//
		//      if (i == 0)
		//      {
		//        firstSpContainer = spc;
		//      }
		//      else
		//      {
		//        .Length += data.Length;
		//      }
		//    }
		//
		//    // Get all the spContainer byte from the charts and add to the array
		//    for (int i = 0 ; i < charts.Length ; i++)
		//    {
		//      EscherContainer spContainer= charts[i].getSpContainer();
		//      byte[] data = spContainer.getBytes();
		//      data = spContainer.setHeaderData(data);
		//      spContainerData[i + numImages] = data;
		//
		//      if (i == 0 && numImages == 0)
		//      {
		//        firstSpContainer = spContainer;
		//      }
		//      else
		//      {
		//        .Length += data.Length;
		//      }
		//    }
		//
		//    // Put the generalised stuff around the first item
		//    Drawing firstDrawing = (Drawing) drawings.get(0);
		//    DgContainer dgContainer = new DgContainer();
		//    Dg dg  = new Dg(numImages + charts.Length);
		//    dgContainer.add(dg);
		//    
		//    SpgrContainer spgrContainer = new SpgrContainer();
		//    
		//    SpContainer spContainer = new SpContainer();
		//    Spgr spgr = new Spgr();
		//    spContainer.add(spgr);
		//    Sp sp = new Sp(Sp.MIN, 1024, 5);
		//    spContainer.add(sp);
		//    spgrContainer.add(spContainer);
		//    
		//    spgrContainer.add(firstSpContainer);
		//    
		//    dgContainer.add(spgrContainer);
		//
		//    byte[] firstMsoData = dgContainer.getData();
		//
		//    // Adjust the .Length of the DgContainer
		//    int len = IntegerHelper.getInt(firstMsoData[4],
		//                                   firstMsoData[5],
		//                                   firstMsoData[6],
		//                                   firstMsoData[7]);
		//    IntegerHelper.getFourBytes(len+.Length, firstMsoData, 4);
		//
		//    // Adjust the .Length of the SpgrContainer
		//    len = IntegerHelper.getInt(firstMsoData[28],
		//                               firstMsoData[29],
		//                               firstMsoData[30],
		//                               firstMsoData[31]);
		//    IntegerHelper.getFourBytes(len+.Length, firstMsoData, 28);
		//
		//    // Now write out each MsoDrawing record and object record
		//
		//    // First MsoRecord
		//    MsoDrawingRecord msoDrawingRecord = new MsoDrawingRecord(firstMsoData);
		//    outputFile.write(msoDrawingRecord);
		//    
		//    ObjRecord objRecord = new ObjRecord(firstDrawing.getObjectId());
		//    outputFile.write(objRecord);
		//
		//    // Now do all the others
		//    for (int i = 1 ; i < spContainerData.Length ; i++)
		//    {
		//      byte[] bytes = (byte[]) spContainerData[i];
		//      msoDrawingRecord = new MsoDrawingRecord(bytes);
		//      outputFile.write(msoDrawingRecord);
		//
		//      if (i < numImages)
		//      {
		//        objRecord = new ObjRecord
		//          (((Drawing) drawings.get(i)).getObjectId());
		//        outputFile.write(objRecord);
		//      }
		//      else
		//      {
		//        Chart chart = charts[i - numImages];
		//        objRecord = chart.getObjRecord();
		//        outputFile.write(objRecord);
		//        outputFile.write(chart);
		//      }
		//    }
		//  }
		//
		
		// [TODO-NExcel_Next]
		//  /**
		//   * Writes out the drawings and the charts if nothing has been modified
		//   *
		//   * @param outputFile
		//   * @exception IOException
		//   */
		//  private void writeUnmodified(File outputFile) throws IOException
		//  {
		//    if (charts.Length == 0 && drawings.size() == 0)
		//    {
		//      // No drawings or charts
		//      return;
		//    }
		//    else if (charts.Length == 0 && drawings.size() != 0)
		//    {
		//      // If there are no charts, then write out the drawings and return
		//      for (Iterator i = drawings.iterator() ; i.hasNext() ; )
		//      {
		//        Drawing d = (Drawing) i.next();
		//        outputFile.write(d.getMsoDrawingRecord());
		//        outputFile.write(d.getObjRecord());
		//      }
		//      return;
		//    }
		//    else if (drawings.size() == 0 && charts.Length != 0)
		//    {
		//      // If there are no drawings, then write out the charts and return
		//      Chart curChart = null;
		//      for (int i = 0 ; i < charts.Length ; i++)
		//      {
		//        curChart = charts[i];
		//        if (curChart.getMsoDrawingRecord() != null)
		//        {
		//          outputFile.write(curChart.getMsoDrawingRecord());
		//        }
		//
		//        if (curChart.getObjRecord() != null)
		//        {
		//          outputFile.write(curChart.getObjRecord());
		//        }
		//
		//        outputFile.write(curChart);
		//      }
		//
		//      return;
		//    }
		//
		//    // There are both charts and drawings - the output 
		//    // drawing group records will need
		//    // to be re-jigged in order to write the drawings out first, then the 
		//    // charts
		//    int numDrawings = drawings.size();
		//    int .Length = 0;
		//    SpContainer[] spContainers = new SpContainer[numDrawings + charts.Length];
		//    
		//    for (int i = 0 ; i < numDrawings; i++)
		//    {
		//      Drawing d = (Drawing) drawings.get(i);
		//      spContainers[i] = d.getSpContainer();
		//
		//      if (i > 0)
		//      {
		//        .Length += spContainers[i].getLength();
		//      }
		//    }
		//
		//    for (int i = 0 ; i < charts.Length ; i++)
		//    {
		//      spContainers[i+numDrawings] = charts[i].getSpContainer();
		//      .Length += spContainers[i+numDrawings].getLength();
		//    }
		//
		//    // Put the generalised stuff around the first item
		//    DgContainer dgContainer = new DgContainer();
		//    Dg dg  = new Dg(numDrawings + charts.Length);
		//    dgContainer.add(dg);
		//    
		//    SpgrContainer spgrContainer = new SpgrContainer();
		//    
		//    SpContainer spContainer = new SpContainer();
		//    Spgr spgr = new Spgr();
		//    spContainer.add(spgr);
		//    Sp sp = new Sp(Sp.MIN, 1024, 5);
		//    spContainer.add(sp);
		//    spgrContainer.add(spContainer);
		//    
		//    spgrContainer.add(spContainers[0]);
		//    
		//    dgContainer.add(spgrContainer);
		//
		//    byte[] firstMsoData = dgContainer.getData();
		//
		//    // Adjust the .Length of the DgContainer
		//    int len = IntegerHelper.getInt(firstMsoData[4],
		//                                   firstMsoData[5],
		//                                   firstMsoData[6],
		//                                   firstMsoData[7]);
		//    IntegerHelper.getFourBytes(len+.Length, firstMsoData, 4);
		//
		//    // Adjust the .Length of the SpgrContainer
		//    len = IntegerHelper.getInt(firstMsoData[28],
		//                               firstMsoData[29],
		//                               firstMsoData[30],
		//                               firstMsoData[31]);
		//    IntegerHelper.getFourBytes(len+.Length, firstMsoData, 28);
		//
		//    // Now write out each MsoDrawing record and object record
		//
		//    // First MsoRecord
		//    MsoDrawingRecord msoDrawingRecord = new MsoDrawingRecord(firstMsoData);
		//    outputFile.write(msoDrawingRecord);
		//    
		//    ObjRecord objRecord = ( (Drawing) drawings.get(0)).getObjRecord();
		//    outputFile.write(objRecord);
		//
		//    // Now do all the others
		//    for (int i = 1 ; i < spContainers.Length ; i++)
		//    {
		//      byte[] bytes = spContainers[i].getBytes();
		//      byte[] bytes2 = spContainers[i].setHeaderData(bytes);
		//      msoDrawingRecord = new MsoDrawingRecord(bytes2);
		//      outputFile.write(msoDrawingRecord);
		//
		//      if (i < numDrawings)
		//      {
		//        objRecord = ((Drawing) drawings.get(i)).getObjRecord();
		//        outputFile.write(objRecord);
		//      }
		//      else
		//      {
		//        Chart chart = charts[i - numDrawings];
		//        objRecord = chart.getObjRecord();
		//        outputFile.write(objRecord);
		//        outputFile.write(chart);
		//      }
		//    }
		//  }
		//
		static SheetDrawingWriter()
		{
			logger = Logger.getLogger(typeof(SheetDrawingWriter));
		}
	}
}
