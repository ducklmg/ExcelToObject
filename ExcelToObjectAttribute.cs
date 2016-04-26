using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToObject
{
	public class ExcelToObjectAttribute : Attribute
	{
		public string TableName { get; private set; }
		public string KeyName { get; private set; }
		public string SheetName { get; private set; }

		public ExcelToObjectAttribute(string TableName = null, string KeyName = null, string SheetName = null)
		{
			this.TableName = TableName;
			this.KeyName = KeyName;
			this.SheetName = SheetName;
		}
	}
}
