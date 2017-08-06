using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToObject
{
	public class ExcelToObjectAttribute : Attribute
	{
		public static readonly ExcelToObjectAttribute Default = new ExcelToObjectAttribute();

		public string TableName { get; set; }
		public string DictionaryKeyName { get; set; }
		public bool Ignore { get; set; }
		public bool MapInto { get; set; }


		public ExcelToObjectAttribute(string tableName = null, string dictionaryKeyName = null, bool ignore = false, bool mapInto = false)
		{
			this.TableName = tableName;
			this.DictionaryKeyName = dictionaryKeyName;
			this.Ignore = ignore;
			this.MapInto = mapInto;
		}
	}
}
