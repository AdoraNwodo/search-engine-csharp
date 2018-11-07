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
using NExcel.Biff;
namespace NExcel.Biff.Drawing
{
	
	class Dgg:EscherAtom
	{
		virtual internal int ShapesSaved
		{
			get
			{
				return shapesSaved;
			}
			
		}
		new private sbyte[] data;
		private int numClusters;
		private int maxShapeId;
		private int shapesSaved;
		private int drawingsSaved;
		
		private ArrayList clusters;
		
		internal sealed class Cluster
		{
			internal int drawingGroupId;
			internal int shapeIdsUsed;
			
			internal Cluster(int dgId, int sids)
			{
				drawingGroupId = dgId;
				shapeIdsUsed = sids;
			}
		}
		
		public Dgg(EscherRecordData erd):base(erd)
		{
			clusters = new ArrayList();
			sbyte[] bytes = Bytes;
			maxShapeId = IntegerHelper.getInt(bytes[0], bytes[1], bytes[2], bytes[3]);
			numClusters = IntegerHelper.getInt(bytes[4], bytes[5], bytes[6], bytes[7]);
			shapesSaved = IntegerHelper.getInt(bytes[8], bytes[9], bytes[10], bytes[11]);
			drawingsSaved = IntegerHelper.getInt(bytes[12], bytes[13], bytes[14], bytes[15]);
			
			int pos = 16;
			for (int i = 0; i < numClusters; i++)
			{
				int dgId = IntegerHelper.getInt(bytes[pos], bytes[pos + 1]);
				int sids = IntegerHelper.getInt(bytes[pos + 2], bytes[pos + 3]);
				Cluster c = new Cluster(dgId, sids);
				clusters.Add(c);
				pos += 4;
			}
		}
		
		public Dgg(int numShapes, int numDrawings):base(EscherRecordType.DGG)
		{
			shapesSaved = numShapes;
			drawingsSaved = numDrawings;
			clusters = new ArrayList();
		}
		
		internal virtual void  addCluster(int dgid, int sids)
		{
			Cluster c = new Cluster(dgid, sids);
			clusters.Add(c);
		}
		
		public override sbyte[] Data
		{
		get
		{
		numClusters = clusters.Count;
		data = new sbyte[16 + numClusters * 4];
		
		// The max shape id
		IntegerHelper.getFourBytes(1024 + shapesSaved, data, 0);
		
		// The number of clusters
		IntegerHelper.getFourBytes(numClusters, data, 4);
		
		// The number of shapes saved
		IntegerHelper.getFourBytes(shapesSaved, data, 8);
		
		// The number of drawings saved
		//    IntegerHelper.getFourBytes(drawingsSaved, data, 12);
		IntegerHelper.getFourBytes(1, data, 12);
		
		int pos = 16;
		for (int i = 0; i < numClusters; i++)
		{
		Cluster c = (Cluster) clusters[i];
		IntegerHelper.getTwoBytes(c.drawingGroupId, data, pos);
		IntegerHelper.getTwoBytes(c.shapeIdsUsed, data, pos + 2);
		pos += 4;
		}
		
		return setHeaderData(data);
		}
		}
		
		internal virtual Cluster getCluster(int i)
		{
			return (Cluster) clusters[i];
		}
	}
}
