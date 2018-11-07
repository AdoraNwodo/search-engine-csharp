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
namespace NExcel.Biff
{
	
	/// <summary> An enumeration class which contains the  biff types</summary>
	public sealed class Type
	{
		/// <summary> The biff value for this type</summary>
		public int Value;
		/// <summary> An array of all types</summary>
		private static NExcel.Biff.Type[] types;
		
		/// <summary> Constructor
		/// Sets the biff value and adds this type to the array of all types
		/// 
		/// </summary>
		/// <param name="v">the biff code for the type
		/// </param>
		private Type(int v)
		{
			Value = v;
			
			// Add to the list of available types
			NExcel.Biff.Type[] newTypes = new NExcel.Biff.Type[types.Length + 1];
			Array.Copy(types, 0, newTypes, 0, types.Length);
			newTypes[types.Length] = this;
			types = newTypes;
		}
		
		/// <summary> Standard hash code method</summary>
		/// <returns> the hash code
		/// </returns>
		public override int GetHashCode()
		{
			return Value;
		}
		
		/// <summary> Standard equals method</summary>
		/// <param name="o">the object to compare
		/// </param>
		/// <returns> TRUE if the objects are equal, FALSE otherwise
		/// </returns>
		public  override bool Equals(System.Object o)
		{
			if (o == this)
			{
				return true;
			}
			
			if (!(o is NExcel.Biff.Type))
			{
				return false;
			}
			
			NExcel.Biff.Type t = (NExcel.Biff.Type) o;
			
			return Value == t.Value;
		}
		
		/// <summary> Gets the type object from its integer value</summary>
		/// <param name="v">the internal code
		/// </param>
		/// <returns> the type
		/// </returns>
		public static NExcel.Biff.Type getType(int v)
		{
			for (int i = 0; i < types.Length; i++)
			{
				if (types[i].Value == v)
				{
					return types[i];
				}
			}
			
			return UNKNOWN;
		}
		
		
		public static readonly NExcel.Biff.Type BOF;
		public static readonly NExcel.Biff.Type EOF;
		public static readonly NExcel.Biff.Type BOUNDSHEET;
		public static readonly NExcel.Biff.Type SUPBOOK;
		public static readonly NExcel.Biff.Type EXTERNSHEET;
		public static readonly NExcel.Biff.Type DIMENSION;
		public static readonly NExcel.Biff.Type BLANK;
		public static readonly NExcel.Biff.Type MULBLANK;
		public static readonly NExcel.Biff.Type ROW;
		public static readonly NExcel.Biff.Type NOTE;
		public static readonly NExcel.Biff.Type TXO;
		public static readonly NExcel.Biff.Type RK;
		public static readonly NExcel.Biff.Type RK2;
		public static readonly NExcel.Biff.Type MULRK;
		public static readonly NExcel.Biff.Type INDEX;
		public static readonly NExcel.Biff.Type DBCELL;
		public static readonly NExcel.Biff.Type SST;
		public static readonly NExcel.Biff.Type COLINFO;
		public static readonly NExcel.Biff.Type EXTSST;
		public static readonly NExcel.Biff.Type CONTINUE;
		public static readonly NExcel.Biff.Type LABEL;
		public static readonly NExcel.Biff.Type RSTRING;
		public static readonly NExcel.Biff.Type LABELSST;
		public static readonly NExcel.Biff.Type NUMBER;
		public static readonly NExcel.Biff.Type NAME;
		public static readonly NExcel.Biff.Type TABID;
		public static readonly NExcel.Biff.Type ARRAY;
		public static readonly NExcel.Biff.Type STRING;
		public static readonly NExcel.Biff.Type FORMULA;
		public static readonly NExcel.Biff.Type FORMULA2;
		public static readonly NExcel.Biff.Type SHAREDFORMULA;
		public static readonly NExcel.Biff.Type FORMAT;
		public static readonly NExcel.Biff.Type XF;
		public static readonly NExcel.Biff.Type BOOLERR;
		public static readonly NExcel.Biff.Type INTERFACEHDR;
		public static readonly NExcel.Biff.Type SAVERECALC;
		public static readonly NExcel.Biff.Type INTERFACEEND;
		public static readonly NExcel.Biff.Type XCT;
		public static readonly NExcel.Biff.Type CRN;
		public static readonly NExcel.Biff.Type DEFCOLWIDTH;
		public static readonly NExcel.Biff.Type DEFAULTROWHEIGHT;
		public static readonly NExcel.Biff.Type WRITEACCESS;
		public static readonly NExcel.Biff.Type WSBOOL;
		public static readonly NExcel.Biff.Type CODEPAGE;
		public static readonly NExcel.Biff.Type DSF;
		public static readonly NExcel.Biff.Type FNGROUPCOUNT;
		public static readonly NExcel.Biff.Type COUNTRY;
		public static readonly NExcel.Biff.Type PROTECT;
		public static readonly NExcel.Biff.Type SCENPROTECT;
		public static readonly NExcel.Biff.Type OBJPROTECT;
		public static readonly NExcel.Biff.Type PRINTHEADERS;
		public static readonly NExcel.Biff.Type HEADER;
		public static readonly NExcel.Biff.Type FOOTER;
		public static readonly NExcel.Biff.Type HCENTER;
		public static readonly NExcel.Biff.Type VCENTER;
		public static readonly NExcel.Biff.Type FILEPASS;
		public static readonly NExcel.Biff.Type SETUP;
		public static readonly NExcel.Biff.Type PRINTGRIDLINES;
		public static readonly NExcel.Biff.Type GRIDSET;
		public static readonly NExcel.Biff.Type GUTS;
		public static readonly NExcel.Biff.Type WINDOWPROTECT;
		public static readonly NExcel.Biff.Type PROT4REV;
		public static readonly NExcel.Biff.Type PROT4REVPASS;
		public static readonly NExcel.Biff.Type PASSWORD;
		public static readonly NExcel.Biff.Type REFRESHALL;
		public static readonly NExcel.Biff.Type WINDOW1;
		public static readonly NExcel.Biff.Type WINDOW2;
		public static readonly NExcel.Biff.Type BACKUP;
		public static readonly NExcel.Biff.Type HIDEOBJ;
		public static readonly NExcel.Biff.Type NINETEENFOUR;
		public static readonly NExcel.Biff.Type PRECISION;
		public static readonly NExcel.Biff.Type BOOKBOOL;
		public static readonly NExcel.Biff.Type FONT;
		public static readonly NExcel.Biff.Type MMS;
		public static readonly NExcel.Biff.Type CALCMODE;
		public static readonly NExcel.Biff.Type CALCCOUNT;
		public static readonly NExcel.Biff.Type REFMODE;
		public static readonly NExcel.Biff.Type TEMPLATE;
		public static readonly NExcel.Biff.Type OBJPROJ;
		public static readonly NExcel.Biff.Type DELTA;
		public static readonly NExcel.Biff.Type MERGEDCELLS;
		public static readonly NExcel.Biff.Type ITERATION;
		public static readonly NExcel.Biff.Type STYLE;
		public static readonly NExcel.Biff.Type USESELFS;
		public static readonly NExcel.Biff.Type HORIZONTALPAGEBREAKS;
		public static readonly NExcel.Biff.Type SELECTION;
		public static readonly NExcel.Biff.Type HLINK;
		public static readonly NExcel.Biff.Type OBJ;
		public static readonly NExcel.Biff.Type MSODRAWING;
		public static readonly NExcel.Biff.Type MSODRAWINGGROUP;
		public static readonly NExcel.Biff.Type LEFTMARGIN;
		public static readonly NExcel.Biff.Type RIGHTMARGIN;
		public static readonly NExcel.Biff.Type TOPMARGIN;
		public static readonly NExcel.Biff.Type BOTTOMMARGIN;
		public static readonly NExcel.Biff.Type EXTERNNAME;
		public static readonly NExcel.Biff.Type PALETTE;
		public static readonly NExcel.Biff.Type PLS;
		public static readonly NExcel.Biff.Type SCL;
		public static readonly NExcel.Biff.Type PANE;
		public static readonly NExcel.Biff.Type WEIRD1;
		public static readonly NExcel.Biff.Type SORT;
		// Chart types
		public static readonly NExcel.Biff.Type FONTX;
		public static readonly NExcel.Biff.Type IFMT;
		public static readonly NExcel.Biff.Type FBI;
		public static readonly NExcel.Biff.Type UNKNOWN;

		static Type()
		{
			types = new NExcel.Biff.Type[0];

			BOF = new NExcel.Biff.Type(0x809);
			EOF = new NExcel.Biff.Type(0x0a);
			BOUNDSHEET = new NExcel.Biff.Type(0x85);
			SUPBOOK = new NExcel.Biff.Type(0x1ae);
			EXTERNSHEET = new NExcel.Biff.Type(0x17);
			DIMENSION = new NExcel.Biff.Type(0x200);
			BLANK = new NExcel.Biff.Type(0x201);
			MULBLANK = new NExcel.Biff.Type(0xbe);
			ROW = new NExcel.Biff.Type(0x208);
			NOTE = new NExcel.Biff.Type(0x1c);
			TXO = new NExcel.Biff.Type(0x1b6);
			RK = new NExcel.Biff.Type(0x7e);
			RK2 = new NExcel.Biff.Type(0x27e);
			MULRK = new NExcel.Biff.Type(0xbd);
			INDEX = new NExcel.Biff.Type(0x20b);
			DBCELL = new NExcel.Biff.Type(0xd7);
			SST = new NExcel.Biff.Type(0xfc);
			COLINFO = new NExcel.Biff.Type(0x7d);
			EXTSST = new NExcel.Biff.Type(0xff);
			CONTINUE = new NExcel.Biff.Type(0x3c);
			LABEL = new NExcel.Biff.Type(0x204);
			RSTRING = new NExcel.Biff.Type(0xd6);
			LABELSST = new NExcel.Biff.Type(0xfd);
			NUMBER = new NExcel.Biff.Type(0x203);
			NAME = new NExcel.Biff.Type(0x18);
			TABID = new NExcel.Biff.Type(0x13d);
			ARRAY = new NExcel.Biff.Type(0x221);
			STRING = new NExcel.Biff.Type(0x207);
			FORMULA = new NExcel.Biff.Type(0x406);
			FORMULA2 = new NExcel.Biff.Type(0x6);
			SHAREDFORMULA = new NExcel.Biff.Type(0x4bc);
			FORMAT = new NExcel.Biff.Type(0x41e);
			XF = new NExcel.Biff.Type(0xe0);
			BOOLERR = new NExcel.Biff.Type(0x205);
			INTERFACEHDR = new NExcel.Biff.Type(0xe1);
			SAVERECALC = new NExcel.Biff.Type(0x5f);
			INTERFACEEND = new NExcel.Biff.Type(0xe2);
			XCT = new NExcel.Biff.Type(0x59);
			CRN = new NExcel.Biff.Type(0x5a);
			DEFCOLWIDTH = new NExcel.Biff.Type(0x55);
			DEFAULTROWHEIGHT = new NExcel.Biff.Type(0x225);
			WRITEACCESS = new NExcel.Biff.Type(0x5c);
			WSBOOL = new NExcel.Biff.Type(0x81);
			CODEPAGE = new NExcel.Biff.Type(0x42);
			DSF = new NExcel.Biff.Type(0x161);
			FNGROUPCOUNT = new NExcel.Biff.Type(0x9c);
			COUNTRY = new NExcel.Biff.Type(0x8c);
			PROTECT = new NExcel.Biff.Type(0x12);
			SCENPROTECT = new NExcel.Biff.Type(0xdd);
			OBJPROTECT = new NExcel.Biff.Type(0x63);
			PRINTHEADERS = new NExcel.Biff.Type(0x2a);
			HEADER = new NExcel.Biff.Type(0x14);
			FOOTER = new NExcel.Biff.Type(0x15);
			HCENTER = new NExcel.Biff.Type(0x83);
			VCENTER = new NExcel.Biff.Type(0x84);
			FILEPASS = new NExcel.Biff.Type(0x2f);
			SETUP = new NExcel.Biff.Type(0xa1);
			PRINTGRIDLINES = new NExcel.Biff.Type(0x2b);
			GRIDSET = new NExcel.Biff.Type(0x82);
			GUTS = new NExcel.Biff.Type(0x80);
			WINDOWPROTECT = new NExcel.Biff.Type(0x19);
			PROT4REV = new NExcel.Biff.Type(0x1af);
			PROT4REVPASS = new NExcel.Biff.Type(0x1bc);
			PASSWORD = new NExcel.Biff.Type(0x13);
			REFRESHALL = new NExcel.Biff.Type(0x1b7);
			WINDOW1 = new NExcel.Biff.Type(0x3d);
			WINDOW2 = new NExcel.Biff.Type(0x23e);
			BACKUP = new NExcel.Biff.Type(0x40);
			HIDEOBJ = new NExcel.Biff.Type(0x8d);
			NINETEENFOUR = new NExcel.Biff.Type(0x22);
			PRECISION = new NExcel.Biff.Type(0xe);
			BOOKBOOL = new NExcel.Biff.Type(0xda);
			FONT = new NExcel.Biff.Type(0x31);
			MMS = new NExcel.Biff.Type(0xc1);
			CALCMODE = new NExcel.Biff.Type(0x0d);
			CALCCOUNT = new NExcel.Biff.Type(0x0c);
			REFMODE = new NExcel.Biff.Type(0x0f);
			TEMPLATE = new NExcel.Biff.Type(0x60);
			OBJPROJ = new NExcel.Biff.Type(0xd3);
			DELTA = new NExcel.Biff.Type(0x10);
			MERGEDCELLS = new NExcel.Biff.Type(0xe5);
			ITERATION = new NExcel.Biff.Type(0x11);
			STYLE = new NExcel.Biff.Type(0x293);
			USESELFS = new NExcel.Biff.Type(0x160);
			HORIZONTALPAGEBREAKS = new NExcel.Biff.Type(0x1b);
			SELECTION = new NExcel.Biff.Type(0x1d);
			HLINK = new NExcel.Biff.Type(0x1b8);
			OBJ = new NExcel.Biff.Type(0x5d);
			MSODRAWING = new NExcel.Biff.Type(0xec);
			MSODRAWINGGROUP = new NExcel.Biff.Type(0xeb);
			LEFTMARGIN = new NExcel.Biff.Type(0x26);
			RIGHTMARGIN = new NExcel.Biff.Type(0x27);
			TOPMARGIN = new NExcel.Biff.Type(0x28);
			BOTTOMMARGIN = new NExcel.Biff.Type(0x29);
			EXTERNNAME = new NExcel.Biff.Type(0x23);
			PALETTE = new NExcel.Biff.Type(0x92);
			PLS = new NExcel.Biff.Type(0x4d);
			SCL = new NExcel.Biff.Type(0xa0);
			PANE = new NExcel.Biff.Type(0x41);
			WEIRD1 = new NExcel.Biff.Type(0xef);
			SORT = new NExcel.Biff.Type(0x90);
			// Chart types
			FONTX = new NExcel.Biff.Type(0x1026);
			IFMT = new NExcel.Biff.Type(0x104e);
			FBI = new NExcel.Biff.Type(0x1060);
			UNKNOWN = new NExcel.Biff.Type(0xffff);

		}
	}
}
