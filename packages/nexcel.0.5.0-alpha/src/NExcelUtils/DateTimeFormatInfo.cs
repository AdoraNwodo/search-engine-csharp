using System;

namespace NExcelUtils
{
	/// <summary>
	/// Custom DateTimeFormatInfo.
	/// Works almost as java.text.SimpleDateFormat.
	/// </summary>
	public class DateTimeFormatInfo : IFormatProvider, ICustomFormatter
	{
		/// <summary>
		/// Pattern.
		/// </summary>
		private string pattern = "dd MM yyyy hh:mm:ss";	// default value


		public DateTimeFormatInfo()
		{
		}


		public DateTimeFormatInfo(string pattern)
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
			if (arg is DateTime)
			{
				return ( (DateTime) arg).ToString(this.pattern);
			}
			// [TODO-NExcel_Next] - add other types

			// otherwise
			return arg.ToString();

		}
	}
}
