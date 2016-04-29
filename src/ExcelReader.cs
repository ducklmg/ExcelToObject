using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ExcelDataReader;

namespace ExcelToObject
{
	public class ExcelReader
	{
		List<SheetData> mSheets;

		public List<SheetData> Sheets
		{
			get
			{
				return mSheets;
			}
		}

		public ExcelReader(byte[] xlsxFile)
		{
			ReadSheetList(xlsxFile);
		}

		public ExcelReader(string filePath)
		{
			ReadSheetList(File.ReadAllBytes(filePath));
		}

		public ExcelReader(Stream stream)
		{
			ReadSheetList(stream.ReadAll());
		}

		void ReadSheetList(byte[] xlsxFile)
		{
			mSheets = ExcelOpenXmlReader.ReadSheets(xlsxFile);
		}


		//////////////////////////////////////////////////////////////////////////
		struct TablePos
		{
			public SheetData Sheet;
			public int Row;
			public int Col;
		}

		TablePos FindTable(string dataName)
		{
			string tag = String.Format("[{0}]", dataName);

			for( int s = 0; s < mSheets.Count; s++ )
			{
				SheetData sheet = mSheets[s];

				for( int r = 0; r < sheet.Rows; r++ )
				{
					for( int c = 0; c < sheet.Columns; c++ )
					{
						if( sheet[r, c] == tag )
						{
							TablePos table;

							table.Sheet = sheet;
							table.Row = r;
							table.Col = c;

							return table;
						}
					}
				}
			}

			return default(TablePos);
		}

		private List<T> ReadListInternal<T>(string dataName, Action<T, string, string> customSetter = null, TableToTypeMap ttm = null, int maxCount = 0) where T : new()
		{
			TablePos table;
			return ReadListInternal<T>(dataName, out table, customSetter, ttm, maxCount);
		}

		private List<T> ReadListInternal<T>(string dataName, out TablePos table, Action<T, string, string> customSetter = null, TableToTypeMap ttm = null, int maxCount = 0) where T : new()
		{
			table = FindTable(dataName);
			if( table.Sheet == null )
				return null;

			var sheet = table.Sheet;
			int rowBase = table.Row + 1;
			int colBase = table.Col;

			// header parse
			if( ttm == null )
				ttm = new TableToTypeMap(typeof(T));

			for( int col = colBase; col < sheet.Columns; col++ )
			{
				string fieldName = sheet[rowBase, col];
				if( fieldName == null )
					break;

				ttm.AddFieldColumn(fieldName);
			}

			List<T> result = new List<T>();

			for( int row = rowBase + 1; row < sheet.Rows; row++ )
			{
				T rowObj = new T();

				ttm.OnNewRow(rowObj);

				bool emptyRow = true;

				for( int col = 0; col < ttm.ColumnCount; col++ )
				{
					string value = sheet[row, colBase + col];
					if( value != null )
					{
						bool ret = ttm.SetValue(rowObj, col, value);

						if( ret == false && customSetter != null )
						{
							customSetter(rowObj, sheet[rowBase, colBase + col], value);
						}

						emptyRow = false;
					}
				}

				if( emptyRow )
					break;

				result.Add(rowObj);

				if( maxCount > 0 && result.Count == maxCount )
					break;
			}

			return result;
		}

		// public methods
		public List<T> ReadList<T>(string dataName, Action<T, string, string> customSetter = null) where T : new()
		{
			return ReadListInternal<T>(dataName, customSetter);
		}

		public T ReadSingle<T>(string dataName, Action<T, string, string> customSetter = null) where T : new()
		{
			var list = ReadListInternal<T>(dataName, customSetter, null, 1);
			return list != null && list.Count >= 1 ? list[0] : default(T);
		}

		public Dictionary<TKey, T> ReadDictionary<TKey, T>(string dataName, Action<T, string, string> customSetter = null) where T : new()
		{
			return ReadDictionary<TKey, T>(dataName, "", customSetter);
		}

		public Dictionary<TKey, T> ReadDictionary<TKey, T>(string dataName, string keyName, Action<T, string, string> customSetter = null) where T : new()
		{
			var ttm = new TableToTypeMap(typeof(T));

			TablePos table;
			List<T> list = ReadListInternal<T>(dataName, out table, customSetter, ttm);
			if( list == null )
				return null;

			var dic = new Dictionary<TKey, T>(list.Count);

			// keyName의 column을 찾는다.
			int col = table.Col;
			int rowBase = table.Row + 1;

			if( keyName.IsValid() )
			{
				for( ; col < table.Sheet.Columns; col++ )
					if( table.Sheet[rowBase, col] == keyName )
						break;

				if( col == table.Sheet.Columns )
					col = table.Col;
			}

			for( int i = 0; i < list.Count; i++ )
			{
				string keyStr = table.Sheet[rowBase + 1 + i, col];

				TKey key = Util.ConvertType<TKey>(keyStr);
				dic[key] = list[i];
			}

			return dic;
		}

		public Dictionary<TKey, T> ReadDictionary<TKey, T>(string dataName, Func<T, TKey> keySelector, Action<T, string, string> customSetter = null) where T : new()
		{
			var ttm = new TableToTypeMap(typeof(T));

			List<T> list = ReadListInternal<T>(dataName, customSetter, ttm);
			if( list == null )
				return null;

			var dic = new Dictionary<TKey, T>(list.Count);

			for( int i = 0; i < list.Count; i++ )
			{
				TKey key = keySelector(list[i]);

				dic[key] = list[i];
			}

			return dic;
		}

		public T ReadValue<T>(string dataName, string propName, T defaultValue = default(T))
		{
			var table = FindTable(dataName);
			if( table.Sheet == null )
				return defaultValue;

			var sheet = table.Sheet;
			int rowBase = table.Row + 1;
			int colBase = table.Col;

			// header parse
			for( int col = colBase; col < sheet.Columns; col++ )
			{
				string fieldName = sheet[rowBase, col];
				if( fieldName == null )
					break;

				if( propName == fieldName )
				{
					string value = sheet[rowBase + 1, col];
					return Util.ConvertType<T>(value);
				}
			}

			return defaultValue;
		}
	}
}
