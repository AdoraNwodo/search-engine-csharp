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
using NExcel.Read.Biff;
namespace NExcel.Biff.Drawing
{
	// [TODO-NExcel_Next]
	//import NExcel.Write.biff.File;
	
	/// <summary> This class contains the Excel picture data in Escher format for the
	/// entire workbook
	/// </summary>
	public class DrawingGroup : EscherStream
	{
		/// <summary> Gets hold of the BStore container from the Escher data
		/// 
		/// </summary>
		/// <returns> the BStore container
		/// </returns>
		private BStoreContainer BStoreContainer
		{
			get
			{
				if (bstoreContainer == null)
				{
					if (!initialized)
					{
						initialize();
					}
					
					EscherRecord[] children = escherData.Children;
					Assert.verify(children[1].Type == EscherRecordType.BSTORE_CONTAINER);
					bstoreContainer = (BStoreContainer) children[1];
				}
				
				return bstoreContainer;
			}
			
		}
		/// <summary> Accessor for the number of blips in the drawing group
		/// 
		/// </summary>
		/// <returns> the number of blips
		/// </returns>
		virtual internal int NumberOfBlips
		{
			get
			{
				return numBlips;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The escher data read in from file</summary>
		private sbyte[] drawingData;
		
		/// <summary> The top level escher container</summary>
		private EscherContainer escherData;
		
		/// <summary> The Bstore container, which contains all the drawing data</summary>
		private BStoreContainer bstoreContainer;
		
		/// <summary> The initialized flag</summary>
		private bool initialized;
		
		/// <summary> The list of user added drawings</summary>
		private ArrayList drawings;
		
		/// <summary> The number of blips</summary>
		private int numBlips;
		
		/// <summary> The number of charts</summary>
		private int numCharts;
		
		/// <summary> The number of shape ids used on the second Dgg cluster</summary>
		private int drawingGroupId;
		
		/// <summary> The origin of this drawing group</summary>
		private Origin origin;
		
		/// <summary> A hash map of images keyed on the file path, containing the
		/// reference count
		/// </summary>
		private Hashtable imageFiles;
		
		public class Origin
		{
		}
		
		public static readonly Origin READ = new Origin();
		public static readonly Origin WRITE = new Origin();
		public static readonly Origin READ_WRITE = new Origin();
		
		/// <summary> Constructor
		/// 
		/// </summary>
		/// <param name="o">the origin of this drawing group
		/// </param>
		public DrawingGroup(Origin o)
		{
			origin = o;
			initialized = o == WRITE?true:false;
			drawings = new ArrayList();
			imageFiles = new Hashtable();
		}
		
		/// <summary> Adds in a drawing group record to this drawing group.  The binary
		/// data is extracted from the drawing group and added to a single
		/// byte array
		/// 
		/// </summary>
		/// <param name="mso">the drawing group record to add
		/// </param>
		public virtual void  add(MsoDrawingGroupRecord mso)
		{
			addData(mso.getData());
		}
		
		public virtual void  add(Record cont)
		{
			addData(cont.Data);
		}
		
		private void  addData(sbyte[] msodata)
		{
			if (drawingData == null)
			{
				drawingData = new sbyte[msodata.Length];
				Array.Copy(msodata, 0, drawingData, 0, msodata.Length);
				return ;
			}
			
			// Grow the array
			sbyte[] newdata = new sbyte[drawingData.Length + msodata.Length];
			Array.Copy(drawingData, 0, newdata, 0, drawingData.Length);
			Array.Copy(msodata, 0, newdata, drawingData.Length, msodata.Length);
			drawingData = newdata;
		}
		
		/// <summary> Adds a drawing to the drawing group
		/// 
		/// </summary>
		/// <param name="d">the drawing to add
		/// </param>
		internal void  addDrawing(Drawing d)
		{
			drawings.Add(d);
		}
		
		/// <summary> Adds a  chart to the darwing group 
		/// 
		/// </summary>
		/// <param name="">c
		/// </param>
		public virtual void  add(Chart c)
		{
			numCharts++;
		}
		
		/// <summary> Adds a drawing from the public, writable interface
		/// 
		/// </summary>
		/// <param name="d">the drawing to add
		/// </param>
		public virtual void  add(Drawing d)
		{
			if (origin == READ)
			{
				origin = READ_WRITE;
				numBlips = BStoreContainer.NumBlips;
				
				Dgg dgg = (Dgg) escherData.Children[0];
				drawingGroupId = dgg.getCluster(1).drawingGroupId - numBlips - 1;
			}
			
			// See if this is referenced elsewhere
			Drawing refImage = (Drawing) imageFiles[d.ImageFilePath];
			
			if (refImage == null)
			{
				// There are no other references to this drawing, so assign
				// a new object id and put it on the hash map
				drawings.Add(d);
				d.DrawingGroup = this;
				d.setObjectId(numBlips + 1, numBlips + 1);
				numBlips++;
				imageFiles[d.ImageFilePath] =  d;
			}
			else
			{
				// This drawing is used elsewhere in the workbook.  Increment the
				// reference count on the drawing, and set the object id of the drawing
				// passed in
				refImage.ReferenceCount = refImage.ReferenceCount + 1;
				d.DrawingGroup = this;
				d.setObjectId(refImage.getObjectId(), refImage.BlipId);
			}
		}
		
		/// <summary> Interface method to remove a drawing from the group
		/// 
		/// </summary>
		/// <param name="d">the drawing to remove
		/// </param>
		public virtual void  remove(Drawing d)
		{
			if (origin == READ)
			{
				origin = READ_WRITE;
				numBlips = BStoreContainer.NumBlips;
				Dgg dgg = (Dgg) escherData.Children[0];
				drawingGroupId = dgg.getCluster(1).drawingGroupId - numBlips - 1;
			}
			
			// Get the blip
			EscherRecord[] children = BStoreContainer.Children;
			BlipStoreEntry bse = (BlipStoreEntry) children[d.BlipId - 1];
			
			bse.dereference();
			
			if (bse.ReferenceCount == 0)
			{
				// Remove the blip
				BStoreContainer.remove(bse);
				
				// Adjust blipId on the other blips
				foreach (Drawing drawing in drawings)
				{
				if (drawing.BlipId > d.BlipId)
				{
				drawing.setObjectId(drawing.getObjectId(), drawing.BlipId - 1);
				}
				}
				
				
				numBlips--;
			}
		}
		
		
		/// <summary> Initializes the drawing data from the escher record read in</summary>
		private void  initialize()
		{
			EscherRecordData er = new EscherRecordData(this, 0);
			
			Assert.verify(er.isContainer());
			
			escherData = new EscherContainer(er);
			
			Assert.verify(escherData.Length == drawingData.Length);
			Assert.verify(escherData.Type == EscherRecordType.DGG_CONTAINER);
			
			initialized = true;
		}
		
		/// <summary> Gets hold of the binary data
		/// 
		/// </summary>
		/// <returns> the data
		/// </returns>
		public virtual sbyte[] getData()
		{
			return drawingData;
		}
		
		// [TODO-NExcel_Next]
		//  /**
		//   * Writes the drawing group to the output file
		//   *
		//   * @param outputFile the file to write to
		//   * @exception IOException
		//   */
		//  public void write(File outputFile) throws IOException
		//  {
		//    if (origin == WRITE)
		//    {
		//      DggContainer dggContainer = new DggContainer();
		//
		//      Dgg dgg = new Dgg(numBlips+numCharts+1, numBlips);
		//
		//      dgg.addCluster(1,0);
		//      dgg.addCluster(numBlips+1,0);
		//
		//      dggContainer.add(dgg);
		//
		//      BStoreContainer bstoreCont = new BStoreContainer(drawings.size());
		//
		//      // Create a blip entry for each drawing
		//      for (Iterator i = drawings.iterator(); i.hasNext();)
		//      {
		//        Drawing d = (Drawing) i.next();
		//        BlipStoreEntry bse = new BlipStoreEntry(d);
		//
		//       bstoreCont.add(bse);
		//      }
		//      dggContainer.add(bstoreCont);
		//
		//      Opt opt = new Opt();
		//
		//      /*
		//      opt.addProperty(191, false, false, 524296);
		//      opt.addProperty(385, false, false, 134217737);
		//      opt.addProperty(448, false, false, 134217792);
		//      */
		//
		//      dggContainer.add(opt);
		//
		//      SplitMenuColors splitMenuColors = new SplitMenuColors();
		//      dggContainer.add(splitMenuColors);
		//
		//      drawingData = dggContainer.getData();
		//    }
		//    else if (origin == READ_WRITE)
		//    {
		//      DggContainer dggContainer = new DggContainer();
		//
		//      Dgg dgg = new Dgg(numBlips+numCharts+1, numBlips);
		//
		//      dgg.addCluster(1,0);
		//      dgg.addCluster(drawingGroupId+numBlips+1,0);
		//
		//      dggContainer.add(dgg);
		//
		//      BStoreContainer bstoreCont = new BStoreContainer(numBlips);
		//
		//      // Create a blip entry for each drawing that was read in
		//      BStoreContainer readBStoreContainer = getBStoreContainer();
		//      EscherRecord[] children = readBStoreContainer.getChildren();
		//      for (int i = 0; i < children.Length ; i++)
		//      {
		//        BlipStoreEntry bse = (BlipStoreEntry) children[i];
		//        bstoreCont.add(bse);
		//      }
		//
		//      // Create a blip entry for each drawing that has been added
		//      for (Iterator i = drawings.iterator(); i.hasNext();)
		//      {
		//        Drawing d = (Drawing) i.next();
		//        if (d.getOrigin() != Drawing.READ)
		//        {
		//          BlipStoreEntry bse = new BlipStoreEntry(d);
		//          bstoreCont.add(bse);
		//        }
		//      }
		//      dggContainer.add(bstoreCont);
		//
		//      Opt opt = new Opt();
		//
		//      opt.addProperty(191, false, false, 524296);
		//      opt.addProperty(385, false, false, 134217737);
		//      opt.addProperty(448, false, false, 134217792);
		//
		//
		//      dggContainer.add(opt);
		//
		//      SplitMenuColors splitMenuColors = new SplitMenuColors();
		//      dggContainer.add(splitMenuColors);
		//
		//      drawingData = dggContainer.getData();
		//
		//    }
		//
		//    MsoDrawingGroupRecord msodg = new MsoDrawingGroupRecord(drawingData);
		//    outputFile.write(msodg);
		//  }
		
		/// <summary> Gets the drawing data for the given blip id.  Called by the Drawing
		/// object
		/// 
		/// </summary>
		/// <param name="blipId">the blipId
		/// </param>
		/// <returns> the drawing data
		/// </returns>
		internal virtual sbyte[] getImageData(int blipId)
		{
			numBlips = BStoreContainer.NumBlips;
			
			Assert.verify(blipId <= numBlips);
			Assert.verify(origin == READ || origin == READ_WRITE);
			
			// Get the blip
			EscherRecord[] children = BStoreContainer.Children;
			BlipStoreEntry bse = (BlipStoreEntry) children[blipId - 1];
			
			return bse.ImageData;
		}
		static DrawingGroup()
		{
			logger = Logger.getLogger(typeof(DrawingGroup));
		}
	}
}
