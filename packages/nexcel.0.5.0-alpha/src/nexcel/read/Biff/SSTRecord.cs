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
namespace NExcel.Read.Biff
{
	
	/// <summary> Holds all the strings in the shared string table</summary>
	public class SSTRecord:RecordData
	{
		/// <summary> The total number of strings in this table</summary>
		private int totalStrings;
		/// <summary> The number of unique strings</summary>
		private int uniqueStrings;
		/// <summary> The shared strings</summary>
		private string[] strings;
		/// <summary> The array of continuation breaks</summary>
		private int[] continuationBreaks;
		
		/// <summary> A holder for a byte array</summary>
		private class ByteArrayHolder
		{
			/// <summary> the byte holder</summary>
			public sbyte[] bytes;
		}
		
		/// <summary> A holder for a boolean</summary>
		private class BooleanHolder
		{
			/// <summary> the holder holder</summary>
			public bool Value;
		}
		
		/// <summary> Constructs this object from the raw data
		/// 
		/// </summary>
		/// <param name="t">the raw data
		/// </param>
		/// <param name="continuations">the continuations
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public SSTRecord(Record t, Record[] continuations, WorkbookSettings ws):base(t)
		{
			
			// If a continue record appears in the middle of
			// a string, then the encoding character is repeated
			
			// Concatenate everything into one big bugger of a byte array
			int totalRecordLength = 0;
			
			for (int i = 0; i < continuations.Length; i++)
			{
				totalRecordLength += continuations[i].Length;
			}
			totalRecordLength += getRecord().Length;
			
			sbyte[] data = new sbyte[totalRecordLength];
			
			// First the original data gets put in
			int pos = 0;
			Array.Copy(getRecord().Data, 0, data, 0, getRecord().Length);
			pos += getRecord().Length;
			
			// Now copy in everything else.
			continuationBreaks = new int[continuations.Length];
			Record r = null;
			for (int i = 0; i < continuations.Length; i++)
			{
				r = continuations[i];
				Array.Copy(r.Data, 0, data, pos, r.Length);
				continuationBreaks[i] = pos;
				pos += r.Length;
			}
			
			totalStrings = IntegerHelper.getInt(data[0], data[1], data[2], data[3]);
			uniqueStrings = IntegerHelper.getInt(data[4], data[5], data[6], data[7]);
			
			strings = new string[uniqueStrings];
			readStrings(data, 8, ws);
		}
		
		/// <summary> Reads in all the strings from the raw data
		/// 
		/// </summary>
		/// <param name="data">the raw data
		/// </param>
		/// <param name="offset">the offset
		/// </param>
		/// <param name="ws">the workbook settings
		/// </param>
		public void  readStrings(sbyte[] data, int offset, WorkbookSettings ws)
		{
			int pos = offset;
			int numChars;
			sbyte optionFlags;
			string s = null;
			bool asciiEncoding = false;
			bool richString = false;
			bool extendedString = false;
			int formattingRuns = 0;
			int extendedRunLength = 0;
			
			for (int i = 0; i < uniqueStrings; i++)
			{
				// Read in the number of characters
				numChars = IntegerHelper.getInt(data[pos], data[pos + 1]);
				pos += 2;
				optionFlags = data[pos];
				pos++;
				
				// See if it is an extended string
				extendedString = ((optionFlags & 0x04) != 0);
				
				// See if string contains formatting information
				richString = ((optionFlags & 0x08) != 0);
				
				if (richString)
				{
					// Read in the crun
					formattingRuns = IntegerHelper.getInt(data[pos], data[pos + 1]);
					pos += 2;
				}
				
				if (extendedString)
				{
					// Read in cchExtRst
					extendedRunLength = IntegerHelper.getInt(data[pos], data[pos + 1], data[pos + 2], data[pos + 3]);
					pos += 4;
				}
				
				// See if string is ASCII (compressed) or unicode
				asciiEncoding = ((optionFlags & 0x01) == 0);
				
				ByteArrayHolder bah = new ByteArrayHolder();
				BooleanHolder bh = new BooleanHolder();
				bh.Value = asciiEncoding;
				pos += getChars(data, bah, pos, bh, numChars);
				asciiEncoding = bh.Value;
				
				if (asciiEncoding)
				{
					s = StringHelper.getString(bah.bytes, numChars, 0, ws);
				}
				else
				{
					s = StringHelper.getUnicodeString(bah.bytes, numChars, 0);
				}
				
				strings[i] = s;
				
				// For rich strings, skip over the formatting runs
				if (richString)
				{
					pos += 4 * formattingRuns;
				}
				
				// For extended strings, skip over the extended string data
				if (extendedString)
				{
					pos += extendedRunLength;
				}
				
				if (pos > data.Length)
				{
					Assert.verify(false, "pos exceeds record .Length");
				}
			}
		}
		
		/// <summary> Gets the chars in the ascii array, taking into account continuation
		/// breaks
		/// 
		/// </summary>
		/// <param name="source">the original source
		/// </param>
		/// <param name="bah">holder for the new byte array
		/// </param>
		/// <param name="pos">the current position in the source
		/// </param>
		/// <param name="ascii">holder for a return ascii flag
		/// </param>
		/// <param name="numChars">the number of chars in the string
		/// </param>
		/// <returns> the number of bytes read from the source
		/// </returns>
		private int getChars(sbyte[] source, ByteArrayHolder bah, int pos, BooleanHolder ascii, int numChars)
		{
			int i = 0;
			bool spansBreak = false;
			//    byte[] bytes = null;
			
			if (ascii.Value)
			{
				bah.bytes = new sbyte[numChars];
			}
			else
			{
				bah.bytes = new sbyte[numChars * 2];
			}
			
			while (i < continuationBreaks.Length && !spansBreak)
			{
				spansBreak = pos <= continuationBreaks[i] && (pos + bah.bytes.Length > continuationBreaks[i]);
				
				if (!spansBreak)
				{
					i++;
				}
			}
			
			// If it doesn't span a break simply do an array copy into the
			// destination array and finish
			if (!spansBreak)
			{
				Array.Copy(source, pos, bah.bytes, 0, bah.bytes.Length);
				return bah.bytes.Length;
			}
			
			// Copy the portion before the break pos into the array
			int breakpos = continuationBreaks[i];
			Array.Copy(source, pos, bah.bytes, 0, breakpos - pos);
			
			int bytesRead = breakpos - pos;
			int charsRead;
			if (ascii.Value)
			{
				charsRead = bytesRead;
			}
			else
			{
				charsRead = bytesRead / 2;
			}
			
			bytesRead += getContinuedString(source, bah, bytesRead, i, ascii, numChars - charsRead);
			return bytesRead;
		}
		
		/// <summary> Gets the rest of the string after a continuation break
		/// 
		/// </summary>
		/// <param name="source">the original bytes
		/// </param>
		/// <param name="bah">the holder for the new bytes
		/// </param>
		/// <param name="destPos">the des pos
		/// </param>
		/// <param name="contBreakIndex">the index of the continuation break
		/// </param>
		/// <param name="ascii">the ascii flag holder
		/// </param>
		/// <param name="charsLeft">the number of chars left in the array
		/// </param>
		/// <returns> the number of bytes read in the continued string
		/// </returns>
		private int getContinuedString(sbyte[] source, ByteArrayHolder bah, int destPos, int contBreakIndex, BooleanHolder ascii, int charsLeft)
		{
			int breakpos = continuationBreaks[contBreakIndex];
			int bytesRead = 0;
			
			while (charsLeft > 0)
			{
				Assert.verify(contBreakIndex < continuationBreaks.Length, "continuation break index");
				
				if (ascii.Value && source[breakpos] == 0)
				{
					// The string is consistently ascii throughout
					
					int length = contBreakIndex == continuationBreaks.Length - 1?charsLeft:System.Math.Min(charsLeft, continuationBreaks[contBreakIndex + 1] - breakpos - 1);
					
					Array.Copy(source, breakpos + 1, bah.bytes, destPos, length);
					destPos += length;
					bytesRead += length + 1;
					charsLeft -= length;
					ascii.Value = true;
				}
				else if (!ascii.Value && source[breakpos] != 0)
				{
					// The string is Unicode throughout
					
					int length = contBreakIndex == continuationBreaks.Length - 1?charsLeft * 2:System.Math.Min(charsLeft * 2, continuationBreaks[contBreakIndex + 1] - breakpos - 1);
					
					// It looks like the string continues as Unicode too.  That's handy
					Array.Copy(source, breakpos + 1, bah.bytes, destPos, length);
					
					destPos += length;
					bytesRead += length + 1;
					charsLeft -= length / 2;
					ascii.Value = false;
				}
				else if (!ascii.Value && source[breakpos] == 0)
				{
					// Bummer - the string starts off as Unicode, but after the
					// continuation it is in straightforward ASCII encoding
					int chars = contBreakIndex == continuationBreaks.Length - 1?charsLeft:System.Math.Min(charsLeft, continuationBreaks[contBreakIndex + 1] - breakpos - 1);
					
					for (int j = 0; j < chars; j++)
					{
						bah.bytes[destPos] = source[breakpos + j + 1];
						destPos += 2;
					}
					
					bytesRead += chars + 1;
					charsLeft -= chars;
					ascii.Value = false;
				}
				else
				{
					// Double Bummer - the string starts off as ASCII, but after the
					// continuation it is in Unicode.  This impacts the allocated array
					
					// Reallocate what we have of the byte array so that it is all
					// Unicode
					sbyte[] oldBytes = bah.bytes;
					bah.bytes = new sbyte[destPos * 2 + charsLeft * 2];
					for (int j = 0; j < destPos; j++)
					{
						bah.bytes[j * 2] = oldBytes[j];
					}
					
					destPos = destPos * 2;
					
					int length = contBreakIndex == continuationBreaks.Length - 1?charsLeft * 2:System.Math.Min(charsLeft * 2, continuationBreaks[contBreakIndex + 1] - breakpos - 1);
					
					Array.Copy(source, breakpos + 1, bah.bytes, destPos, length);
					
					destPos += length;
					bytesRead += length + 1;
					charsLeft -= length / 2;
					ascii.Value = false;
				}
				
				contBreakIndex++;
				
				if (contBreakIndex < continuationBreaks.Length)
				{
					breakpos = continuationBreaks[contBreakIndex];
				}
			}
			
			return bytesRead;
		}
		
		/// <summary> Gets the string at the specified position
		/// 
		/// </summary>
		/// <param name="index">the index of the string to return
		/// </param>
		/// <returns> the strings
		/// </returns>
		public virtual string getString(int index)
		{
			Assert.verify(index < uniqueStrings);
			return strings[index];
		}
	}
}
