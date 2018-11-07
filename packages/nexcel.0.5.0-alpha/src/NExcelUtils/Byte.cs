using System;

namespace NExcelUtils
{
	/// <summary>
	/// Byte and SByte utils.
	/// </summary>
	internal class Byte
	{
		private Byte()
		{
		}


		/// <summary>
		/// Converts an array of sbytes to an array of bytes
		/// </summary>
		/// <param name="sbyteArray">The array of sbytes to be converted</param>
		/// <returns>The new array of bytes</returns>
		public static byte[] ToByteArray(sbyte[] sbyteArray)
		{
			byte[] byteArray = new byte[sbyteArray.Length];
			for(int index=0; index < sbyteArray.Length; index++)
				byteArray[index] = (byte) sbyteArray[index];
			return byteArray;
		}


		/// <summary>
		/// Converts a string to an array of bytes
		/// </summary>
		/// <param name="sourceString">The string to be converted</param>
		/// <returns>The new array of bytes</returns>
		public static byte[] ToByteArray(string sourceString)
		{
			byte[] byteArray = new byte[sourceString.Length];
			for (int index=0; index < sourceString.Length; index++)
				byteArray[index] = (byte) sourceString[index];
			return byteArray;
		}

		/// <summary>
		/// Converts a array of object-type instances to a byte-type array.
		/// </summary>
		/// <param name="tempObjectArray">Array to convert.</param>
		/// <returns>An array of byte type elements.</returns>
		public static byte[] ToByteArray(object[] tempObjectArray)
		{
			byte[] byteArray = new byte[tempObjectArray.Length];
			for (int index = 0; index < tempObjectArray.Length; index++)
				byteArray[index] = (byte)tempObjectArray[index];
			return byteArray;
		}

		/// <summary>
		/// Receives a byte array and returns it transformed in an sbyte array
		/// </summary>
		/// <param name="byteArray">Byte array to process</param>
		/// <returns>The transformed array</returns>
		public static sbyte[] ToSByteArray(byte[] byteArray)
		{
			sbyte[] sbyteArray = new sbyte[byteArray.Length];
			for(int index=0; index < byteArray.Length; index++)
				sbyteArray[index] = (sbyte) byteArray[index];
			return sbyteArray;
		}



		/// <summary>
		/// Converts an array of sbytes to an array of chars
		/// </summary>
		/// <param name="sByteArray">The array of sbytes to convert</param>
		/// <returns>The new array of chars</returns>
		public static char[] ToCharArray(sbyte[] sByteArray) 
		{
			char[] charArray = new char[sByteArray.Length];	   
			sByteArray.CopyTo(charArray, 0);
			return charArray;
		}

		/// <summary>
		/// Converts an array of bytes to an array of chars
		/// </summary>
		/// <param name="byteArray">The array of bytes to convert</param>
		/// <returns>The new array of chars</returns>
		public static char[] ToCharArray(byte[] byteArray) 
		{
			char[] charArray = new char[byteArray.Length];	   
			byteArray.CopyTo(charArray, 0);
			return charArray;
		}

	}
}
