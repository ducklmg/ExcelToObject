using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToObject
{
	public class SheetData
	{
		private string[,] Data;

		public string Name { get; private set; }
		public int Rows { get; private set; }
		public int Columns { get; private set; }

		public SheetData(string name, string[,] data)
		{
			Name = name;
			Data = data;
			Rows = data.GetLength(0);
			Columns = data.GetLength(1);
		}

		public string this[int row, int col]
		{
			get
			{
				return Data[row, col];
			}
		}
	}
}
