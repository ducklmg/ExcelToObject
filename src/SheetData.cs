using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToObject
{
	public class SheetData
	{
		string mName;
		string[,] mData;
		int mRows;
		int mColumns;

		public string Name { get { return mName; } }
		public string[,] Data { get { return mData; } }
		public int Rows { get { return mRows; } }
		public int Columns { get { return mColumns; } }

		public SheetData(string name, string[,] data)
		{
			mName = name;
			mData = data;
			mRows = data.GetLength(0);
			mColumns = data.GetLength(1);
		}

		public string this[int row, int col]
		{
			get
			{
				return mData[row, col];
			}
		}

		public Table FindTable(string name)
		{
			for( int r = 0; r < mRows; r++ )
			{
				for( int c = 0; c < mColumns; c++ )
				{
					string value = mData[r, c];

					if( Table.IsTableMarker(value, name) )
					{
						return new Table(mData, r, c);
					}
				}
			}

			return null;
		}

		public void FindTables(string name, List<Table> result)
		{
			for( int r = 0; r < mRows; r++ )
			{
				for( int c = 0; c < mColumns; c++ )
				{
					string value = mData[r, c];

					if( Table.IsTableMarker(value, name) )
					{
						result.Add(new Table(mData, r, c));
					}
				}
			}
		}

		public List<Table> GetAllTables()
		{
			var result = new List<Table>();

			FindTables(null, result);

			return result;
		}
	}
}
