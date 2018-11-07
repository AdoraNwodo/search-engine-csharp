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
namespace NExcel.Biff
{
	
	/// <summary> Contains the common data for a compound file</summary>
	public abstract class BaseCompoundFile
	{
		/// <summary> The identifier at the beginning of every OLE file</summary>
		protected internal static readonly sbyte[] IDENTIFIER = new sbyte[] {
																			   -0x30, 
																			   -0x31, 
																			   0x11, 
																			   -0x20, 
																			   -0x5F, 
																			   -0x4F, 
																			   0x1a, 
																			   -0x1F,
																		   };
		
		protected internal const int NUM_BIG_BLOCK_DEPOT_BLOCKS_POS = 0x2c;
		
		protected internal const int SMALL_BLOCK_DEPOT_BLOCK_POS = 0x3c;
		
		protected internal const int ROOT_START_BLOCK_POS = 0x30;
		
		protected internal const int BIG_BLOCK_SIZE = 0x200;
		
		protected internal const int SMALL_BLOCK_SIZE = 0x40;
		
		protected internal const int EXTENSION_BLOCK_POS = 0x44;
		
		protected internal const int NUM_EXTENSION_BLOCK_POS = 0x48;
		
		protected internal const int PROPERTY_STORAGE_BLOCK_SIZE = 0x80;
		
		protected internal const int BIG_BLOCK_DEPOT_BLOCKS_POS = 0x4c;
		
		protected internal const int SMALL_BLOCK_THRESHOLD = 0x1000;
		
		// property storage offsets
		
		private const int SIZE_OF_NAME_POS = 0x40;
		
		private const int TYPE_POS = 0x42;
		
		private const int PREVIOUS_POS = 0x44;
		
		private const int NEXT_POS = 0x48;
		
		private const int DIRECTORY_POS = 0x4c;
		
		private const int START_BLOCK_POS = 0x74;
		
		private const int SIZE_POS = 0x78;
		
		/// <summary> Inner class to represent the property storage sets</summary>
		protected internal class PropertyStorage
		{
			private void  InitBlock(BaseCompoundFile enclosingInstance)
			{
				this.enclosingInstance = enclosingInstance;
			}
			private BaseCompoundFile enclosingInstance;
			/// <summary> Sets the type
			/// 
			/// </summary>
			/// <param name="t">the type
			/// </param>
			virtual public int Type
			{
				set
				{
					type = value;
					data[NExcel.Biff.BaseCompoundFile.TYPE_POS] = (sbyte) value;
					data[NExcel.Biff.BaseCompoundFile.TYPE_POS + 1] = (sbyte) (0x1);
				}
				
			}
			/// <summary> Sets the number of the start block
			/// 
			/// </summary>
			/// <param name="sb">the number of the start block
			/// </param>
			virtual public int StartBlock
			{
				set
				{
					startBlock = value;
					IntegerHelper.getFourBytes(value, data, NExcel.Biff.BaseCompoundFile.START_BLOCK_POS);
				}
				
			}
			/// <summary> Sets the size of the file
			/// 
			/// </summary>
			/// <param name="s">the size
			/// </param>
			virtual public int Size
			{
				set
				{
					size = value;
					IntegerHelper.getFourBytes(value, data, NExcel.Biff.BaseCompoundFile.SIZE_POS);
				}
				
			}
			/// <summary> Sets the previous block
			/// 
			/// </summary>
			/// <param name="prev">the previous block
			/// </param>
			virtual public int Previous
			{
				set
				{
					previous = value;
					IntegerHelper.getFourBytes(value, data, NExcel.Biff.BaseCompoundFile.PREVIOUS_POS);
				}
				
			}
			/// <summary> Sets the next block
			/// 
			/// </summary>
			/// <param name="nxt">the next block
			/// </param>
			virtual public int Next
			{
				set
				{
					next = value;
					IntegerHelper.getFourBytes(next, data, NExcel.Biff.BaseCompoundFile.NEXT_POS);
				}
				
			}
			/// <summary> Sets the directory
			/// 
			/// </summary>
			/// <param name="dir">the directory
			/// </param>
			virtual public int Directory
			{
				set
				{
					directory = value;
					IntegerHelper.getFourBytes(directory, data, NExcel.Biff.BaseCompoundFile.DIRECTORY_POS);
				}
				
			}
			public BaseCompoundFile Enclosing_Instance
			{
				get
				{
					return enclosingInstance;
				}
				
			}
			/// <summary> The name of this property set</summary>
			public string name;
			/// <summary> The type of the property set</summary>
			public int type;
			/// <summary> The block number in the stream which this property sets starts at</summary>
			public int startBlock;
			/// <summary> The size, in bytes, of this property set</summary>
			public int size;
			/// <summary> The previous property set</summary>
			public int previous;
			/// <summary> The next property set</summary>
			public int next;
			/// <summary> The directory for this property set</summary>
			public int directory;
			
			/// <summary> The data that created this set</summary>
			public sbyte[] data;
			
			/// <summary> Constructs a property set
			/// 
			/// </summary>
			/// <param name="d">the bytes
			/// </param>
			public PropertyStorage(BaseCompoundFile enclosingInstance, sbyte[] d)
			{
				InitBlock(enclosingInstance);
				data = d;
				int nameSize = IntegerHelper.getInt(data[NExcel.Biff.BaseCompoundFile.SIZE_OF_NAME_POS], data[NExcel.Biff.BaseCompoundFile.SIZE_OF_NAME_POS + 1]);
				type = data[NExcel.Biff.BaseCompoundFile.TYPE_POS];
				
				startBlock = IntegerHelper.getInt(data[NExcel.Biff.BaseCompoundFile.START_BLOCK_POS], data[NExcel.Biff.BaseCompoundFile.START_BLOCK_POS + 1], data[NExcel.Biff.BaseCompoundFile.START_BLOCK_POS + 2], data[NExcel.Biff.BaseCompoundFile.START_BLOCK_POS + 3]);
				size = IntegerHelper.getInt(data[NExcel.Biff.BaseCompoundFile.SIZE_POS], data[NExcel.Biff.BaseCompoundFile.SIZE_POS + 1], data[NExcel.Biff.BaseCompoundFile.SIZE_POS + 2], data[NExcel.Biff.BaseCompoundFile.SIZE_POS + 3]);
				
				int chars = 0;
				if (nameSize > 2)
				{
					chars = (nameSize - 1) / 2;
				}
				
				System.Text.StringBuilder n = new System.Text.StringBuilder("");
				for (int i = 0; i < chars; i++)
				{
					n.Append((char) data[i * 2]);
				}
				
				name = n.ToString();
			}
			
			/// <summary> Constructs an empty property set.  Used when writing the file
			/// 
			/// </summary>
			/// <param name="name">the property storage name
			/// </param>
			public PropertyStorage(BaseCompoundFile enclosingInstance, string name)
			{
				InitBlock(enclosingInstance);
				data = new sbyte[NExcel.Biff.BaseCompoundFile.PROPERTY_STORAGE_BLOCK_SIZE];
				
				Assert.verify(name.Length < 32);
				
				IntegerHelper.getTwoBytes((name.Length + 1) * 2, data, NExcel.Biff.BaseCompoundFile.SIZE_OF_NAME_POS);
				// add one to the name .Length to allow for the null character at
				// the end
				for (int i = 0; i < name.Length; i++)
				{
					data[i * 2] = (sbyte) name[i];
				}
			}
		}
		
		/// <summary> Constructor</summary>
		protected internal BaseCompoundFile()
		{
		}
	}
}
