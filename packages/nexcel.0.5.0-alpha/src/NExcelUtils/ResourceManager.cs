using System;
using System.Globalization;
using System.Reflection;
using System.Resources;

namespace NExcelUtils
{
	/// <summary>
	/// Custom ResourceManager per globalzation.
	/// It loads resources with a name_xx, where xx is the language of the input culture.
	/// Doesn't use .NET handling for international resources.
	/// The differencee is that all resources are inside the assembly, and there 
	/// are no satellite assemblies.
	/// </summary>
	internal class ResourceManager
	{
		System.Resources.ResourceManager rm;


		public ResourceManager(string name, CultureInfo culture, Assembly assembly)
		{
			this.rm = new System.Resources.ResourceManager(name + this.GetSuffix(culture), assembly);
		}


		public string GetString(string key)
		{
			return this.rm.GetString(key);
		}


		/// <summary>
		/// Returns the 2-chars culture.
		/// </summary>
		/// <param name="culture"></param>
		/// <returns></returns>
		private string GetSuffix(CultureInfo culture)
		{
			string name = culture.Name;
			if (name==null) return "";
			if (name.Length<2) return "";
			string lang = name.Substring(0,2).ToLower();
			if (lang=="de") return "_de";
			if (lang=="es") return "_es";
			if (lang=="fr") return "_fr";
			
			// otherwise
			return "";
		}

	}
}
