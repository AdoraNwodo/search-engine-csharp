using System;

namespace NExcelUtils
{
	/// <summary>
	/// Custom NumberFormatInfo.
	/// Works almost as java.text.NumberFormat.
	/// </summary>
	public class NumberFormatInfo : IFormatProvider, ICustomFormatter
	{
		/// <summary>
		/// Pattern.
		/// </summary>
		private string pattern = "0";	// default value

		public string Pattern
		{
			get
			{
				return this.pattern;
			}
		}

		public NumberFormatInfo()
		{
		}


		public NumberFormatInfo(string pattern)
		{
			this.pattern = pattern;
		}


		// String.Format calls this method to get an instance of an
		// ICustomFormatter to handle the formatting.
		public object GetFormat (Type service)
		{
			if (service == typeof (ICustomFormatter)) 
			{
				return this;
			}
			else
			{
				return null;
			}
		}
		// After String.Format gets the ICustomFormatter, it calls this format
		// method on each argument.
		public string Format (string format, object arg, IFormatProvider provider) 
		{
			if (arg is double)
			{
				return ( (double) arg).ToString(this.pattern);
			}
			if (arg is int)
			{
				return ( (int) arg).ToString(this.pattern);
			}
			if (arg is long)
			{
				return ( (long) arg).ToString(this.pattern);
			}
			if (arg is decimal)
			{
				return ( (decimal) arg).ToString(this.pattern);
			}
			// [TODO-NExcel_Next] - add other types

			// otherwise
			return arg.ToString();

		}
	
	}
}
