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
namespace NExcel.Read.Biff
{
	
	/// <summary> Reads in and defrags an OLE compound compound file
	/// (Made public only for the PropertySets demo)
	/// </summary>
	public sealed class CompoundFile:BaseCompoundFile
	{
		/// <summary> Gets the property sets</summary>
		/// <returns> the list of property sets
		/// </returns>
		public string[] PropertySetNames
		{
			get
			{
				string[] sets = new string[propertySets.Count];
				for (int i = 0; i < sets.Length; i++)
				{
					PropertyStorage ps = (PropertyStorage) propertySets[i];
					sets[i] = ps.name;
				}
				
				return sets;
			}
			
		}
		/// <summary> The logger</summary>
		private static Logger logger;
		
		/// <summary> The original OLE stream, organized into blocks, which can
		/// appear at any physical location in the file
		/// </summary>
		private sbyte[] data;
		/// <summary> The number of blocks it takes to store the big block depot</summary>
		private int numBigBlockDepotBlocks;
		/// <summary> The start block of the small block depot</summary>
		private int sbdStartBlock;
		/// <summary> The start block of the root entry</summary>
		private int rootStartBlock;
		/// <summary> The header extension block</summary>
		private int extensionBlock;
		/// <summary> The number of header extension blocks</summary>
		private int numExtensionBlocks;
		/// <summary> The root entry</summary>
		private sbyte[] rootEntry;
		/// <summary> The sequence of blocks which comprise the big block chain</summary>
		private int[] bigBlockChain;
		/// <summary> The sequence of blocks which comprise the small block chain</summary>
		private int[] smallBlockChain;
		/// <summary> The chain of blocks which comprise the big block depot</summary>
		private int[] bigBlockDepotBlocks;
		/// <summary> The list of property sets</summary>
		private ArrayList propertySets;
		
		/// <summary> The workbook settings</summary>
		private WorkbookSettings settings;
		
		/// <summary> Initializes the compound file
		/// 
		/// </summary>
		/// <param name="d">the raw data of the ole stream
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		/// <exception cref=""> BiffException
		/// </exception>
		public CompoundFile(sbyte[] d, WorkbookSettings ws):base()
		{
			data = d;
			settings = ws;
			
			// First verify the OLE identifier
			for (int i = 0; i < IDENTIFIER.Length; i++)
			{
				if (data[i] != IDENTIFIER[i])
				{
					throw new BiffException(BiffException.unrecognizedOLEFile);
				}
			}
			
			propertySets = new ArrayList();
			numBigBlockDepotBlocks = IntegerHelper.getInt(data[NUM_BIG_BLOCK_DEPOT_BLOCKS_POS], data[NUM_BIG_BLOCK_DEPOT_BLOCKS_POS + 1], data[NUM_BIG_BLOCK_DEPOT_BLOCKS_POS + 2], data[NUM_BIG_BLOCK_DEPOT_BLOCKS_POS + 3]);
			sbdStartBlock = IntegerHelper.getInt(data[SMALL_BLOCK_DEPOT_BLOCK_POS], data[SMALL_BLOCK_DEPOT_BLOCK_POS + 1], data[SMALL_BLOCK_DEPOT_BLOCK_POS + 2], data[SMALL_BLOCK_DEPOT_BLOCK_POS + 3]);
			rootStartBlock = IntegerHelper.getInt(data[ROOT_START_BLOCK_POS], data[ROOT_START_BLOCK_POS + 1], data[ROOT_START_BLOCK_POS + 2], data[ROOT_START_BLOCK_POS + 3]);
			extensionBlock = IntegerHelper.getInt(data[EXTENSION_BLOCK_POS], data[EXTENSION_BLOCK_POS + 1], data[EXTENSION_BLOCK_POS + 2], data[EXTENSION_BLOCK_POS + 3]);
			numExtensionBlocks = IntegerHelper.getInt(data[NUM_EXTENSION_BLOCK_POS], data[NUM_EXTENSION_BLOCK_POS + 1], data[NUM_EXTENSION_BLOCK_POS + 2], data[NUM_EXTENSION_BLOCK_POS + 3]);
			
			bigBlockDepotBlocks = new int[numBigBlockDepotBlocks];
			
			int pos = BIG_BLOCK_DEPOT_BLOCKS_POS;
			
			int bbdBlocks = numBigBlockDepotBlocks;
			
			if (numExtensionBlocks != 0)
			{
				bbdBlocks = (BIG_BLOCK_SIZE - BIG_BLOCK_DEPOT_BLOCKS_POS) / 4;
			}
			
			for (int i = 0; i < bbdBlocks; i++)
			{
				bigBlockDepotBlocks[i] = IntegerHelper.getInt(d[pos], d[pos + 1], d[pos + 2], d[pos + 3]);
				pos += 4;
			}
			
			for (int j = 0; j < numExtensionBlocks; j++)
			{
				pos = (extensionBlock + 1) * BIG_BLOCK_SIZE;
				int blocksToRead = System.Math.Min(numBigBlockDepotBlocks - bbdBlocks, BIG_BLOCK_SIZE / 4 - 1);
				
				for (int i = bbdBlocks; i < bbdBlocks + blocksToRead; i++)
				{
					bigBlockDepotBlocks[i] = IntegerHelper.getInt(d[pos], d[pos + 1], d[pos + 2], d[pos + 3]);
					pos += 4;
				}
				
				bbdBlocks += blocksToRead;
				if (bbdBlocks < numBigBlockDepotBlocks)
				{
					extensionBlock = IntegerHelper.getInt(d[pos], d[pos + 1], d[pos + 2], d[pos + 3]);
				}
			}
			
			readBigBlockDepot();
			readSmallBlockDepot();
			
			rootEntry = readData(rootStartBlock);
			readPropertySets();
		}
		
		/// <summary> Reads the big block depot entries</summary>
		private void  readBigBlockDepot()
		{
			int pos = 0;
			int index = 0;
			bigBlockChain = new int[numBigBlockDepotBlocks * BIG_BLOCK_SIZE / 4];
			
			for (int i = 0; i < numBigBlockDepotBlocks; i++)
			{
				pos = (bigBlockDepotBlocks[i] + 1) * BIG_BLOCK_SIZE;
				
				for (int j = 0; j < BIG_BLOCK_SIZE / 4; j++)
				{
					bigBlockChain[index] = IntegerHelper.getInt(data[pos], data[pos + 1], data[pos + 2], data[pos + 3]);
					pos += 4;
					index++;
				}
			}
		}
		
		/// <summary> Reads the small block depot entries</summary>
		private void  readSmallBlockDepot()
		{
			int pos = 0;
			int index = 0;
			int sbdBlock = sbdStartBlock;
			smallBlockChain = new int[0];
			
			while (sbdBlock != - 2)
			{
				// Allocate some more space to the small block chain
				int[] oldChain = smallBlockChain;
				smallBlockChain = new int[smallBlockChain.Length + BIG_BLOCK_SIZE / 4];
				Array.Copy(oldChain, 0, smallBlockChain, 0, oldChain.Length);
				
				pos = (sbdBlock + 1) * BIG_BLOCK_SIZE;
				
				for (int j = 0; j < BIG_BLOCK_SIZE / 4; j++)
				{
					smallBlockChain[index] = IntegerHelper.getInt(data[pos], data[pos + 1], data[pos + 2], data[pos + 3]);
					pos += 4;
					index++;
				}
				
				sbdBlock = bigBlockChain[sbdBlock];
			}
		}
		
		/// <summary> Reads all the property sets</summary>
		private void  readPropertySets()
		{
			int offset = 0;
			sbyte[] d = null;
			
			while (offset < rootEntry.Length)
			{
				d = new sbyte[PROPERTY_STORAGE_BLOCK_SIZE];
				Array.Copy(rootEntry, offset, d, 0, d.Length);
				PropertyStorage ps = new PropertyStorage(this, d);
				propertySets.Add(ps);
				offset += PROPERTY_STORAGE_BLOCK_SIZE;
			}
		}
		
		/// <summary> Gets the defragmented stream from this ole compound file
		/// 
		/// </summary>
		/// <param name="streamName">the stream name to get
		/// </param>
		/// <returns> the defragmented ole stream
		/// </returns>
		/// <exception cref=""> BiffException
		/// </exception>
		public sbyte[] getStream(string streamName)
		{
			PropertyStorage ps = getPropertyStorage(streamName);
			
			if (ps.size >= SMALL_BLOCK_THRESHOLD || streamName.ToUpper().Equals("root entry".ToUpper()))
			{
				return getBigBlockStream(ps);
			}
			else
			{
				return getSmallBlockStream(ps);
			}
		}
		
		/// <summary> Gets the property set with the specified name</summary>
		/// <param name="name">the property storage name
		/// </param>
		/// <returns> the property storage record
		/// </returns>
		/// <exception cref=""> BiffException
		/// </exception>
		private PropertyStorage getPropertyStorage(string name)
		{
			// Find the workbook property
			bool found = false;
			PropertyStorage psres = null;
			foreach(PropertyStorage ps in propertySets)
			{
				if (found) break;
				if (ps.name.ToLower().Equals(name==null ? null : name.ToLower()))
				{
					found = true;
					psres = ps;
				}
			}
			
			if (!found)
			{
				throw new BiffException(BiffException.streamNotFound);
			}
			
			return psres;
		}
		
		/// <summary> Build up the resultant stream using the big blocks
		/// 
		/// </summary>
		/// <param name="ps">the property storage
		/// </param>
		/// <returns> the big block stream
		/// </returns>
		private sbyte[] getBigBlockStream(PropertyStorage ps)
		{
			int numBlocks = ps.size / BIG_BLOCK_SIZE;
			if (ps.size % BIG_BLOCK_SIZE != 0)
			{
				numBlocks++;
			}
			
			sbyte[] streamData = new sbyte[numBlocks * BIG_BLOCK_SIZE];
			
			int block = ps.startBlock;
			
			int count = 0;
			int pos = 0;
			while (block != - 2 && count < numBlocks)
			{
				pos = (block + 1) * BIG_BLOCK_SIZE;
				Array.Copy(data, pos, streamData, count * BIG_BLOCK_SIZE, BIG_BLOCK_SIZE);
				count++;
				block = bigBlockChain[block];
			}
			
			if (block != - 2 && count == numBlocks)
			{
				logger.warn("Property storage size inconsistent with block chain.");
			}
			
			return streamData;
		}
		
		/// <summary> Build up the resultant stream using the small blocks</summary>
		/// <param name="ps">the property storage
		/// </param>
		/// <returns>  the data
		/// </returns>
		/// <exception cref=""> BiffException
		/// </exception>
		private sbyte[] getSmallBlockStream(PropertyStorage ps)
		{
			PropertyStorage rootps = null;
			try
			{
				rootps = getPropertyStorage("root entry");
			}
			catch (BiffException e)
			{
				rootps = (PropertyStorage) propertySets[0];
			}
			
			sbyte[] rootdata = readData(rootps.startBlock);
			sbyte[] sbdata = new sbyte[0];
			
			int block = ps.startBlock;
			//    int count = 0;
			int pos = 0;
			while (block != - 2)
			{
				// grow the array
				sbyte[] olddata = sbdata;
				sbdata = new sbyte[olddata.Length + SMALL_BLOCK_SIZE];
				Array.Copy(olddata, 0, sbdata, 0, olddata.Length);
				
				// Copy in the new data
				pos = block * SMALL_BLOCK_SIZE;
				Array.Copy(rootdata, pos, sbdata, olddata.Length, SMALL_BLOCK_SIZE);
				block = smallBlockChain[block];
			}
			
			return sbdata;
		}
		
		/// <summary> Reads the block chain from the specified block and returns the
		/// data as a continuous stream of bytes
		/// </summary>
		/// <param name="bl">the block number
		/// </param>
		/// <returns> the data
		/// </returns>
		private sbyte[] readData(int bl)
		{
			int block = bl;
			int pos = 0;
			sbyte[] entry = new sbyte[0];
			
			while (block != - 2)
			{
				// Grow the array
				sbyte[] oldEntry = entry;
				entry = new sbyte[oldEntry.Length + BIG_BLOCK_SIZE];
				Array.Copy(oldEntry, 0, entry, 0, oldEntry.Length);
				pos = (block + 1) * BIG_BLOCK_SIZE;
				Array.Copy(data, pos, entry, oldEntry.Length, BIG_BLOCK_SIZE);
				block = bigBlockChain[block];
			}
			return entry;
		}
		static CompoundFile()
		{
			logger = Logger.getLogger(typeof(CompoundFile));
		}
	}
}
