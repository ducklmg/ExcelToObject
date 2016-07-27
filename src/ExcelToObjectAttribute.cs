using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToObject
{
	public class ExcelToObjectAttribute : Attribute
	{
		public string TablePath { get; set; }
		public string DictionaryKeyName { get; set; }
		public bool Ignore { get; set; }

		public ExcelToObjectAttribute(string tablePath = null, string dictionaryKeyName = null, bool ignore = false)
		{
			this.TablePath = tablePath;
			this.DictionaryKeyName = dictionaryKeyName;
			this.Ignore = ignore;
		}
	}
}
