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

namespace NExcel.Biff.Drawing
{
	
	/// <summary> Contains the various biff records used to insert a drawing into a 
	/// worksheet
	/// </summary>
	public class Drawing : EscherStream
	{
		/// <summary> Accessor for the image file
		/// 
		/// </summary>
		/// <returns> the image file
		/// </returns>
		virtual protected internal System.IO.FileInfo ImageFile
		{
			get
			{
				return imageFile;
			}
			
		}
		/// <summary> Accessor for the image file path.  Normally this is the absolute path
		/// of a file on the directory system, but if this drawing was constructed
		/// using an byte[] then the blip id is returned
		/// 
		/// </summary>
		/// <returns> the image file path, or the blip id
		/// </returns>
		virtual protected internal string ImageFilePath
		{
			get
			{
				if (imageFile == null)
				{
					// return the blip id, if it exists
					return blipId != 0?System.Convert.ToString(blipId):"__new__image__";
				}
				
				return imageFile.FullName;
			}
			
		}
		/// <summary> Accessor for the blip id
		/// 
		/// </summary>
		/// <returns> the blip id
		/// </returns>
		virtual internal int BlipId
		{
			get
			{
				if (!initialized)
				{
					initialize();
				}
				
				return blipId;
			}
			
		}
		/// <summary> Gets the drawing record which was read in
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
		/// <summary> Gets the obj record which was read in
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
		/// <summary> Creates the main Sp container for the drawing
		/// 
		/// </summary>
		/// <returns> the SP container
		/// </returns>
		virtual public SpContainer SpContainer
		{
			get
			{
				if (!initialized)
				{
					initialize();
				}
				
				if (origin == READ)
				{
					return ReadSpContainer;
				}
				
				SpContainer spContainer = new SpContainer();
				Sp sp = new Sp(Sp.PICTURE_FRAME, 1024 + objectId, 2560);
				spContainer.add(sp);
				Opt opt = new Opt();
				opt.addProperty(260, true, false, blipId);
				string filePath = imageFile != null?imageFile.FullName:"";
				opt.addProperty(261, true, true, filePath.Length * 2, filePath);
				opt.addProperty(447, false, false, 65536);
				opt.addProperty(959, false, false, 524288);
				spContainer.add(opt);
				ClientAnchor clientAnchor = new ClientAnchor(x, y, x + width, y + height);
				spContainer.add(clientAnchor);
				ClientData clientData = new ClientData();
				spContainer.add(clientData);
				
				return spContainer;
			}
			
		}
		/// <summary> Accessor for the drawing group
		/// 
		/// </summary>
		/// <returns> the drawing group
		/// </returns>
		/// <summary> Sets the drawing group for this drawing.  Called by the drawing group
		/// when this drawing is added to it
		/// 
		/// </summary>
		/// <param name="dg">the drawing group
		/// </param>
		virtual internal DrawingGroup DrawingGroup
		{
			get
			{
				return drawingGroup;
			}
			
			set
			{
				drawingGroup = value;
			}
			
		}
		/// <summary> Accessor for the reference count on this drawing
		/// 
		/// </summary>
		/// <returns> the reference count
		/// </returns>
		/// <summary> Sets the new reference count on the drawing
		/// 
		/// </summary>
		/// <param name="r">the new reference count
		/// </param>
		virtual internal int ReferenceCount
		{
			get
			{
				return referenceCount;
			}
			
			set
			{
				referenceCount = value;
			}
			
		}
		/// <summary> Accessor for the column of this drawing
		/// 
		/// </summary>
		/// <returns> the column
		/// </returns>
		/// <summary> Sets the column position of this drawing
		/// 
		/// </summary>
		/// <param name="x">the column
		/// </param>
		virtual public double X
		{
			get
			{
				if (!initialized)
				{
					initialize();
				}
				return x;
			}
			
			set
			{
				if (origin == READ)
				{
					if (!initialized)
					{
						initialize();
					}
					origin = READ_WRITE;
				}
				
				this.x = value;
			}
			
		}
		/// <summary> Accessor for the row of this drawing
		/// 
		/// </summary>
		/// <returns> the row
		/// </returns>
		/// <summary> Accessor for the row of the drawing
		/// 
		/// </summary>
		/// <param name="y">the row
		/// </param>
		virtual public double Y
		{
			get
			{
				if (!initialized)
				{
					initialize();
				}
				
				return y;
			}
			
			set
			{
				if (origin == READ)
				{
					if (!initialized)
					{
						initialize();
					}
					origin = READ_WRITE;
				}
				
				this.y = value;
			}
			
		}
		/// <summary> Accessor for the width of this drawing
		/// 
		/// </summary>
		/// <returns> the number of columns spanned by this image
		/// </returns>
		/// <summary> Accessor for the width
		/// 
		/// </summary>
		/// <param name="w">the number of columns to span
		/// </param>
		virtual public double Width
		{
			get
			{
				if (!initialized)
				{
					initialize();
				}
				
				return width;
			}
			
			set
			{
				if (origin == READ)
				{
					if (!initialized)
					{
						initialize();
					}
					origin = READ_WRITE;
				}
				
				width = value;
			}
			
		}
		/// <summary> Accessor for the height of this drawing
		/// 
		/// </summary>
		/// <returns> the number of rows spanned by this image
		/// </returns>
		/// <summary> Accessor for the height of this drawing
		/// 
		/// </summary>
		/// <param name="h">the number of rows spanned by this image
		/// </param>
		virtual public double Height
		{
			get
			{
				if (!initialized)
				{
					initialize();
				}
				
				return height;
			}
			
			set
			{
				if (origin == READ)
				{
					if (!initialized)
					{
						initialize();
					}
					origin = READ_WRITE;
				}
				
				height = value;
			}
			
		}
		/// <summary> Gets the SpContainer that was read in
		/// 
		/// </summary>
		/// <returns> the read sp container
		/// </returns>
		private SpContainer ReadSpContainer
		{
			get
			{
				if (!initialized)
				{
					initialize();
				}
				
				return readSpContainer;
			}
			
		}
		/// <summary> Accessor for the image data
		/// 
		/// </summary>
		/// <returns> the image data
		/// </returns>
		virtual public sbyte[] ImageData
		{
			get
			{
				Assert.verify(origin == READ || origin == READ_WRITE);
				
				if (!initialized)
				{
					initialize();
				}
				
				return drawingGroup.getImageData(blipId);
			}
			
		}
		/// <summary> Accessor for the image data
		/// 
		/// </summary>
		/// <returns> the image data
		/// </returns>
		virtual internal sbyte[] ImageBytes
		{
			get
			{
				if (origin == READ || origin == READ_WRITE)
				{
					return ImageData;
				}
				
				Assert.verify(origin == WRITE);
				
				if (imageFile == null)
				{
					Assert.verify(imageData != null);
					return imageData;
				}
				
				sbyte[] data = new sbyte[imageFile.Length];
				System.IO.FileStream fis = new System.IO.FileStream(imageFile.FullName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
				File.ReadInput(fis, ref data, 0, data.Length);
				fis.Close();
				return data;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The entire  drawing data read in</summary>
		private sbyte[] drawingData;
		
		/// <summary> The spContainer that was read in</summary>
		private SpContainer readSpContainer;
		
		/// <summary> The MsoDrawingRecord associated with the drawing</summary>
		private MsoDrawingRecord msoDrawingRecord;
		
		/// <summary> The ObjRecord associated with the drawing</summary>
		private ObjRecord objRecord;
		
		/// <summary> Initialized flag</summary>
		private bool initialized = false;
		
		/// <summary> The file containing the image</summary>
		private System.IO.FileInfo imageFile;
		
		/// <summary> The raw image data, used instead of an image file</summary>
		private sbyte[] imageData;
		
		/// <summary> The object id, assigned by the drawing group</summary>
		private int objectId;
		
		/// <summary> The blip id</summary>
		private int blipId;
		
		/// <summary> The column position of the image</summary>
		private double x;
		
		/// <summary> The row position of the image</summary>
		private double y;
		
		/// <summary> The width of the image in cells</summary>
		private double width;
		
		/// <summary> The height of the image in cells</summary>
		private double height;
		
		/// <summary> The number of places this drawing is referenced</summary>
		private int referenceCount;
		
		/// <summary> The top level escher container</summary>
		private EscherContainer escherData;
		
		/// <summary> Where this image came from (read, written or a copy)</summary>
		private Origin origin;
		
		/// <summary> The drawing group for all the images</summary>
		private DrawingGroup drawingGroup;
		
		// Enumerations for the origin
		public sealed class Origin
		{
		}
		
		public static readonly Origin READ = new Origin();
		public static readonly Origin WRITE = new Origin();
		public static readonly Origin READ_WRITE = new Origin();
		
		/// <summary> Constructor used when reading images
		/// 
		/// </summary>
		/// <param name="mso">the drawing record
		/// </param>
		/// <param name="obj">the object record
		/// </param>
		/// <param name="dg">the drawing group
		/// </param>
		public Drawing(MsoDrawingRecord mso, ObjRecord obj, DrawingGroup dg)
		{
			drawingGroup = dg;
			msoDrawingRecord = mso;
			objRecord = obj;
			initialized = false;
			origin = READ;
			drawingData = msoDrawingRecord.getData();
			
			Assert.verify(mso != null && obj != null);
			
			initialize();
			
			if (blipId != 0)
			{
				drawingGroup.addDrawing(this);
			}
			else
			{
				logger.warn("linked drawings are not supported");
			}
		}
		
		/// <summary> Copy constructor used to copy drawings from read to write
		/// 
		/// </summary>
		/// <param name="d">the drawing to copy
		/// </param>
		protected internal Drawing(Drawing d)
		{
			Assert.verify(d.origin == READ);
			msoDrawingRecord = d.msoDrawingRecord;
			objRecord = d.objRecord;
			initialized = false;
			origin = READ;
			drawingData = d.drawingData;
			drawingGroup = d.drawingGroup;
		}
		
		/// <summary> Constructor invoked when writing the images
		/// 
		/// </summary>
		/// <param name="x">the column
		/// </param>
		/// <param name="y">the row
		/// </param>
		/// <param name="width">the width in cells
		/// </param>
		/// <param name="height">the height in cells
		/// </param>
		/// <param name="image">the image file
		/// </param>
		public Drawing(double x, double y, double width, double height, System.IO.FileInfo image)
		{
			imageFile = image;
			initialized = true;
			origin = WRITE;
			this.x = x;
			this.y = y;
			this.width = width;
			this.height = height;
			referenceCount = 1;
		}
		
		/// <summary> Constructor invoked when writing the images
		/// 
		/// </summary>
		/// <param name="x">the column
		/// </param>
		/// <param name="y">the row
		/// </param>
		/// <param name="width">the width in cells
		/// </param>
		/// <param name="height">the height in cells
		/// </param>
		/// <param name="image">the image data
		/// </param>
		public Drawing(double x, double y, double width, double height, sbyte[] image)
		{
			imageData = image;
			initialized = true;
			origin = WRITE;
			this.x = x;
			this.y = y;
			this.width = width;
			this.height = height;
			referenceCount = 1;
		}
		
		/// <summary> Initializes the member variables from the Escher stream data</summary>
		private void  initialize()
		{
			EscherRecordData er = new EscherRecordData(this, 0);
			Assert.verify(er.isContainer());
			
			escherData = new EscherContainer(er);
			
			readSpContainer = null;
			if (escherData.Type == EscherRecordType.DG_CONTAINER)
			{
				EscherRecordData erd = new EscherRecordData(this, 80);
				Assert.verify(erd.Type == EscherRecordType.SP_CONTAINER);
				readSpContainer = new SpContainer(erd);
			}
			else
			{
				Assert.verify(escherData.Type == EscherRecordType.SP_CONTAINER);
				readSpContainer = new SpContainer(er);
			}
			
			Sp sp = (Sp) readSpContainer.Children[0];
			objectId = sp.ShapeId - 1024;
			
			Opt opt = (Opt) readSpContainer.Children[1];
			
			if (opt.getProperty(260) != null)
			{
				blipId = opt.getProperty(260).Value;
			}
			
			if (opt.getProperty(261) != null)
			{
				imageFile = new System.IO.FileInfo(opt.getProperty(261).stringValue);
			}
			else
			{
				logger.warn("no filename property for drawing");
				imageFile = new System.IO.FileInfo(System.Convert.ToString(blipId));
			}
			
			ClientAnchor clientAnchor = (ClientAnchor) readSpContainer.Children[2];
			x = clientAnchor.X1;
			y = clientAnchor.Y1;
			width = clientAnchor.X2 - x;
			height = clientAnchor.Y2 - y;
			
			initialized = true;
		}
		
		/// <summary> Sets the object id.  Invoked by the drawing group when the object is 
		/// added to id
		/// 
		/// </summary>
		/// <param name="objid">the object id
		/// </param>
		/// <param name="bip">the blip id
		/// </param>
		internal void  setObjectId(int objid, int bip)
		{
			objectId = objid;
			blipId = bip;
			
			if (origin == READ)
			{
				origin = READ_WRITE;
			}
		}
		
		/// <summary> Accessor for the object id
		/// 
		/// </summary>
		/// <returns> the object id
		/// </returns>
		internal int getObjectId()
		{
			return objectId;
		}
		
		/// <summary> Gets the data which was read in for this drawing
		/// 
		/// </summary>
		/// <returns> the drawing data
		/// </returns>
		public virtual sbyte[] getData()
		{
			return drawingData;
		}
		
		/// <summary> Gets the origin of this drawing
		/// 
		/// </summary>
		/// <returns> where this drawing came from
		/// </returns>
		internal virtual Origin getOrigin()
		{
			return origin;
		}
		static Drawing()
		{
			logger = Logger.getLogger(typeof(Drawing));
		}
	}
}
