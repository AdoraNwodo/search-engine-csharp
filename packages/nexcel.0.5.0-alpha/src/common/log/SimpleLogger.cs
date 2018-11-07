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
using Logger = common.Logger;
namespace common.log
{
	
	/// <summary> The default logger.  Simple writes everything out to stdout or stderr</summary>
	public class SimpleLogger:Logger
	{
		/// <summary> Overrides the method in the base class to suppress warnings - it can
		/// be set using the system property NExcel.nowarnings.  
		/// This method was originally present in the WorkbookSettings bean,
		/// but has been moved to the logger class.  This means it is now present
		/// when the JVM is initialized, and subsequent to change it on 
		/// a Workbook by Workbook basis will prove fruitless
		/// 
		/// </summary>
		/// <param name="w">suppression flag
		/// </param>
		override public bool SuppressWarnings
		{
			set
			{
				suppressWarnings = value;
			}
			
		}
		/// <summary> Flag to indicate whether or not warnings should be suppressed</summary>
		private bool suppressWarnings;
		
		/// <summary> Constructor</summary>
		public SimpleLogger()
		{
			suppressWarnings = false;
		}
		
		/// <summary>  Log a debug message</summary>
		public override void  debug(System.Object message)
		{
			if (!suppressWarnings)
			{
				System.Console.Out.Write("Debug: ");
				System.Console.Out.WriteLine(message);
			}
		}
		
		/// <summary> Log a debug message and exception</summary>
		public override void  debug(System.Object message, System.Exception t)
		{
			if (!suppressWarnings)
			{
				System.Console.Out.Write("Debug: ");
				System.Console.Out.WriteLine(message);
			}
		}
		
		/// <summary>  Log an error message</summary>
		public override void  error(System.Object message)
		{
			System.Console.Error.Write("Error: ");
			System.Console.Error.WriteLine(message);
		}
		
		/// <summary> Log an error message object and exception</summary>
		public override void  error(System.Object message, System.Exception t)
		{
			System.Console.Error.Write("Error: ");
			System.Console.Error.WriteLine(message);
			//    t.printStackTrace();
		}
		
		/// <summary> Log a fatal message</summary>
		public override void  fatal(System.Object message)
		{
			System.Console.Error.Write("Fatal: ");
			System.Console.Error.WriteLine(message);
		}
		
		/// <summary> Log a fatal message and exception</summary>
		public override void  fatal(System.Object message, System.Exception t)
		{
			System.Console.Error.Write("Fatal:  ");
			System.Console.Error.WriteLine(message);
			//    t.printStackTrace();
		}
		
		/// <summary> Log an information message</summary>
		public override void  info(System.Object message)
		{
			if (!suppressWarnings)
			{
				System.Console.Out.WriteLine(message);
			}
		}
		
		/// <summary> Logs an information message and an exception</summary>
		
		public override void  info(System.Object message, System.Exception t)
		{
			if (!suppressWarnings)
			{
				System.Console.Out.WriteLine(message);
				//      t.printStackTrace();
			}
		}
		
		/// <summary> Log a warning message object</summary>
		public override void  warn(System.Object message)
		{
			if (!suppressWarnings)
			{
				System.Console.Error.Write("Warning:  ");
				System.Console.Error.WriteLine(message);
			}
		}
		
		/// <summary> Log a warning message with exception</summary>
		public override void  warn(System.Object message, System.Exception t)
		{
			if (!suppressWarnings)
			{
				System.Console.Error.Write("Warning:  ");
				System.Console.Error.WriteLine(message);
				//      t.printStackTrace();
			}
		}
		
		/// <summary> Accessor to the logger implementation</summary>
		protected internal override Logger getLoggerImpl(System.Type c)
		{
			return this;
		}
	}
}