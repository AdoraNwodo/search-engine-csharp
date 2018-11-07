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
namespace NExcel.Biff.Drawing
{
	
	/// <summary> An escher container.  This record may contain other escher containers or
	/// atoms
	/// </summary>
	public class EscherContainer:EscherRecord
	{
		virtual public EscherRecord[] Children
		{
			get
			{
				if (!initialized)
				{
					initialize();
				}
				
				System.Object[] ca = children.ToArray();
				EscherRecord[] era = new EscherRecord[ca.Length];
				Array.Copy(ca, 0, era, 0, ca.Length);
				
				return era;
			}
			
		}
		override public sbyte[] Data
		{
			get
			{
				sbyte[] data = new sbyte[0];
				foreach (EscherRecord er in children)
				{
				sbyte[] childData = er.Data;
				sbyte[] newData = new sbyte[data.Length + childData.Length];
				Array.Copy(data, 0, newData, 0, data.Length);
				Array.Copy(childData, 0, newData, data.Length, childData.Length);
				data = newData;
				}
				
				return setHeaderData(data);
			}
			
		}
		private bool initialized;
		private ArrayList children;
		
		public EscherContainer(EscherRecordData erd):base(erd)
		{
			initialized = false;
			children = new ArrayList();
		}
		
		protected internal EscherContainer(EscherRecordType type):base(type)
		{
			Container = true;
			children = new ArrayList();
		}
		
		public virtual void  add(EscherRecord child)
		{
			children.Add(child);
		}
		
		public virtual void  remove(EscherRecord child)
		{
			children.Remove(child);
		}
		
		private void  initialize()
		{
			int curpos = Pos + HEADER_LENGTH;
			int endpos = Pos + Length;
			
			EscherRecord newRecord = null;
			
			while (curpos < endpos)
			{
				EscherRecordData erd = new EscherRecordData(EscherStream, curpos);
				
				EscherRecordType type = erd.Type;
				if (type == EscherRecordType.DGG)
				{
					newRecord = new Dgg(erd);
				}
				else if (type == EscherRecordType.DG)
				{
					newRecord = new Dg(erd);
				}
				else if (type == EscherRecordType.BSTORE_CONTAINER)
				{
					newRecord = new BStoreContainer(erd);
				}
				else if (type == EscherRecordType.SPGR_CONTAINER)
				{
					newRecord = new SpgrContainer(erd);
				}
				else if (type == EscherRecordType.SP_CONTAINER)
				{
					newRecord = new SpContainer(erd);
				}
				else if (type == EscherRecordType.SPGR)
				{
					newRecord = new Spgr(erd);
				}
				else if (type == EscherRecordType.SP)
				{
					newRecord = new Sp(erd);
				}
				else if (type == EscherRecordType.CLIENT_ANCHOR)
				{
					newRecord = new ClientAnchor(erd);
				}
				else if (type == EscherRecordType.CLIENT_DATA)
				{
					newRecord = new ClientData(erd);
				}
				else if (type == EscherRecordType.BSE)
				{
					newRecord = new BlipStoreEntry(erd);
				}
				else if (type == EscherRecordType.OPT)
				{
					newRecord = new Opt(erd);
				}
				else if (type == EscherRecordType.SPLIT_MENU_COLORS)
				{
					newRecord = new SplitMenuColors(erd);
				}
				else
				{
					newRecord = new EscherAtom(erd);
				}
				
				children.Add(newRecord);
				curpos += newRecord.Length;
			}
			
			initialized = true;
		}
	}
}
