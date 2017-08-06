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

		public Table FindTable(string name)
		{
			string tag = String.Format("[{0}]", name);

			for( int r = 0; r < Rows; r++ )
			{
				for( int c = 0; c < Columns; c++ )
				{
					if( Data[r, c] == tag )
					{
						return new Table(name, this, r, c);
					}
				}
			}

			return null;
		}
	}
}
