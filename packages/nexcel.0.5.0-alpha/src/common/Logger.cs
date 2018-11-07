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
namespace common
{
	
	/// <summary> Abstract wrapper class for the logging interface of choice.  
	/// The methods declared here are the same as those for the log4j  
	/// </summary>
	public abstract class Logger
	{
		/// <summary> Empty implementation of the suppressWarnings.  Subclasses may 
		/// or may not override this method.  This method is included
		/// primarily for backwards support of the NExcel.nowarnings property, and
		/// is used only by the SimpleLogger
		/// 
		/// </summary>
		/// <param name="w">suppression flag
		/// </param>
		virtual public bool SuppressWarnings
		{
			set
			{
				// default implementation does nothing
			}
			
		}
		/// <summary> The singleton logger</summary>
		private static Logger logger = null;
		
		/// <summary> Factory method to return the logger</summary>
		public static Logger getLogger(System.Type cl)
		{
			if (logger == null)
			{
				initializeLogger();
			}
			
			return logger.getLoggerImpl(cl);
		}
		
		/// <summary> Initializes the logger in a thread safe manner</summary>
		private static void  initializeLogger()
		{
			lock (typeof(common.Logger))
			{
				if (logger != null)
				{
					return ;
				}
				
				// First see if there was anything defined at run time
				logger = new common.log.SimpleLogger();
			}
		}
		
		/// <summary> Constructor</summary>
		protected internal Logger()
		{
		}
		
		/// <summary>  Log a debug message</summary>
		public abstract void  debug(System.Object message);
		
		/// <summary> Log a debug message and exception</summary>
		public abstract void  debug(System.Object message, System.Exception t);
		
		/// <summary>  Log an error message</summary>
		public abstract void  error(System.Object message);
		
		/// <summary> Log an error message object and exception</summary>
		public abstract void  error(System.Object message, System.Exception t);
		
		/// <summary> Log a fatal message</summary>
		public abstract void  fatal(System.Object message);
		
		/// <summary> Log a fatal message and exception</summary>
		public abstract void  fatal(System.Object message, System.Exception t);
		
		/// <summary> Log an information message</summary>
		public abstract void  info(System.Object message);
		
		/// <summary> Logs an information message and an exception</summary>
		public abstract void  info(System.Object message, System.Exception t);
		
		/// <summary> Log a warning message object</summary>
		public abstract void  warn(System.Object message);
		
		/// <summary> Log a warning message with exception</summary>
		public abstract void  warn(System.Object message, System.Exception t);
		
		/// <summary> Accessor to the logger implementation</summary>
		protected internal abstract Logger getLoggerImpl(System.Type cl);
	}
}