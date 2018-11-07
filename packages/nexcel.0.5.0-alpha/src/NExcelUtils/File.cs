using System;
using System.IO;

namespace NExcelUtils
{
	/// <summary>
	/// File utils.
	/// </summary>
	internal class File
	{
		private File()
		{
		}

		/// <summary>Reads a number of characters from the current source Stream and writes the data to the target array at the specified index.</summary>
		/// <param name="sourceStream">The source Stream to read from.</param>
		/// <param name="target">Contains the array of characteres read from the source Stream.</param>
		/// <param name="start">The starting index of the target array.</param>
		/// <param name="count">The maximum number of characters to read from the source Stream.</param>
		/// <returns>The number of characters read. The number will be less than or equal to count depending on the data available in the source Stream. Returns -1 if the end of the stream is reached.</returns>
		public static int ReadInput(Stream sourceStream, ref sbyte[] target, int start, int count)
		{
			// Returns 0 bytes if not enough space in target
			if (target.Length == 0)
				return 0;

			byte[] receiver = new byte[target.Length];
			int bytesRead   = sourceStream.Read(receiver, start, count);

			// Returns -1 if EOF
			if (bytesRead == 0)	
				return -1;
                
			for(int i = start; i < start + bytesRead; i++)
				target[i] = (sbyte)receiver[i];
                
			return bytesRead;
		}

	}
}
