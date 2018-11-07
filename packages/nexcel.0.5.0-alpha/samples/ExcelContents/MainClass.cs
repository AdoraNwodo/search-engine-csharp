using System;
using System.IO;

namespace ExcelContents
{
	/// <summary>
	/// Main class.
	/// </summary>
	class MainClass
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main(string[] args)
		{
			// read input parameters
			Param param = new Param(args);
			// if input parameters are not valid, writes usage and exits
			if (param.IsValid()==false)
			{
				WriteUsage();
				return;
			}

			// reads the Excel file and writes contents to console
			ExcelEngine engine = new ExcelEngine();
			engine.Write(param.Filename);

		}


		#region "Write Utils"

		/// <summary>
		/// Writes to console the usage text.
		/// </summary>
		private static void WriteUsage()
		{
			Console.WriteLine("Writes to console the contents of a Excel file.");
			Console.WriteLine("Usage: ExcelContents  filename");
			Console.WriteLine("");
			Console.WriteLine("  filename             the Excel file name");
		}

		#endregion

		#region "Param Class"

		/// <summary>
		/// Container class of application's input parameters.
		/// </summary>
		class Param
		{
			/// <summary>
			/// File name.
			/// Otherwise is null.
			/// </summary>
			public string Filename = null;

			/// <summary>
			/// Create a new instance, reading data from args.
			/// </summary>
			/// <param name="args">the application's input parameters</param>
			public Param(string[] args)
			{
				// check input parameters
				if (args==null) return;
				if (args.Length<1) return;

				// set input parameters
				try
				{
					FileInfo fi = new FileInfo(args[0]);
					this.Filename = fi.FullName;
				}
				catch
				{
					this.Filename = null;
				}
			}

			/// <summary>
			/// Returns true if it has valid data.
			/// Otherwise returns false.
			/// </summary>
			/// <returns></returns>
			public bool IsValid()
			{
				if (this.Filename==null) return false;

				return true;
			}
		}
		#endregion

	}
}
