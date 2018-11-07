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
using NExcel.Biff;
namespace NExcel.Biff.Drawing
{
	
	/// <summary> An options record in the escher stream</summary>
	class Opt:EscherAtom
	{
		/// <summary> The logger</summary>
		private static Logger logger;
		
		new private sbyte[] data;
		private int numProperties;
		
		/// <summary> The list of properties</summary>
		private ArrayList properties;
		
		internal sealed class Property
		{
			internal int id;
			internal bool blipId;
			internal bool complex;
			internal int Value;
			internal string stringValue;
			
			public Property(int i, bool bl, bool co, int v)
			{
				id = i;
				blipId = bl;
				complex = co;
				Value = v;
			}
			
			public Property(int i, bool bl, bool co, int v, string s)
			{
				id = i;
				blipId = bl;
				complex = co;
				Value = v;
				stringValue = s;
			}
		}
		
		public Opt(EscherRecordData erd):base(erd)
		{
			numProperties = Instance;
			readProperties();
		}
		
		private void  readProperties()
		{
			properties = new ArrayList();
			int pos = 0;
			sbyte[] bytes = Bytes;
			
			for (int i = 0; i < numProperties; i++)
			{
				int val = IntegerHelper.getInt(bytes[pos], bytes[pos + 1]);
				int id = val & 0x3fff;
				int Value = IntegerHelper.getInt(bytes[pos + 2], bytes[pos + 3], bytes[pos + 4], bytes[pos + 5]);
				Property p = new Property(id, (val & 0x4000) != 0, (val & 0x8000) != 0, Value);
				pos += 6;
				properties.Add(p);
			}
			
			foreach (Property p in properties)
			{
			if (p.complex)
			{
			p.stringValue = StringHelper.getUnicodeString(bytes, p.Value/2,
			pos);
			pos += p.Value;
			}
			}
		}
		
		public Opt():base(EscherRecordType.OPT)
		{
			properties = new ArrayList();
			Version = 3;
		}
		
		public override sbyte[] Data
		{
		get
		{
		numProperties = properties.Count;
		Instance = numProperties;
		
		data = new sbyte[numProperties * 6];
		int pos = 0;
		
		// Add in the root data
		foreach (Property p in properties)
		{
		int val = p.id & 0x3fff;
		
		if (p.blipId)
		{
		val |= 0x4000;
		}
		
		if (p.complex)
		{
		val |= 0x8000;
		}
		
		IntegerHelper.getTwoBytes(val, data, pos);
		IntegerHelper.getFourBytes(p.Value, data, pos+2);
		pos += 6 ;
		}
		
		// Add in any complex data
		foreach (Property p in properties)
		{
		if (p.complex && p.stringValue != null)
		{
		sbyte[] newData = 
		new sbyte[data.Length + p.stringValue.Length * 2];
		System.Array.Copy(data, 0, newData, 0, data.Length);
		StringHelper.getUnicodeBytes(p.stringValue, newData, data.Length);
		data = newData;
		}
		}
		
		return setHeaderData(data);
		}
		}
		
		
		internal virtual void  addProperty(int id, bool blip, bool complex, int val)
		{
			Property p = new Property(id, blip, complex, val);
			properties.Add(p);
		}
		
		internal virtual void  addProperty(int id, bool blip, bool complex, int val, string s)
		{
			Property p = new Property(id, blip, complex, val, s);
			properties.Add(p);
		}
		
		internal virtual Property getProperty(int id)
		{
			Property pres = null;
			
			foreach (Property p in properties)
			{
			if (p.id == id)
			{
			pres = p;
			break;
			}
			}
			
			return pres;
			
		}
		static Opt()
		{
			logger = Logger.getLogger(typeof(Opt));
		}
	}
}
