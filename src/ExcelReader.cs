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
		public Table FindTable(string name)
		{
			for( int i = 0; i < mSheets.Count; i++ )
			{
				var table = mSheets[i].FindTable(name);
				if( table != null )
					return table;
			}

			return null;
		}

		class ReadResult<T>
		{
			public TableToTypeMap ttm;
			public TablePos tablePos;
		}

		delegate bool CustomParser<T>(T obj, string name, string value);

		private List<T> ReadListInternal<T>(string dataName, CustomParser<T> customParser = null, int maxCount = 0, ReadResult result=null) where T : new()
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

						if( ret == false && customParser != null )
						{
							customParser(rowObj, sheet[rowBase, colBase + col], value);
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
