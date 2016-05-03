using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ExcelDataReader;

namespace ExcelToObject
{
	/// <summary>
	/// Read and parse excel table.
	/// </summary>
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

		/// <summary>
		/// Find table with given name.
		/// In excel, table should be marked with braketed name ([table-name])
		/// </summary>
		/// <param name="name">Table name (without braket)</param>
		/// <returns>Table instance. If not found, returns null.</returns>
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

		/// <summary>
		/// Custom parser signiture for individual properties.
		/// </summary>
		/// <typeparam name="T">Instance type</typeparam>
		/// <param name="obj">instance</param>
		/// <param name="name">property name (column)</param>
		/// <param name="value">property value. null if empty cell.</param>
		/// <returns></returns>
		public delegate bool CustomParser<T>(T obj, string name, string value);


		//////////////////////////////////////////////////////////////
		class ReadContext<T>
		{
			public int readMaxRows;
			public CustomParser<T> customParser;

			public Table table;
			public TableToTypeMap ttm;
		}

		private List<T> ReadListInternal<T>(string tableName, ReadContext<T> context) where T : new()
		{
			// find table
			Table table = FindTable(tableName);
			if( table == null )
				return null;

			// header parse
			var ttm = new TableToTypeMap(typeof(T));

			for( int col = 0; col < table.Columns; col++ )
			{
				ttm.AddFieldColumn(table.GetColumnName(col));
			}

			// rows parse
			List<T> result = new List<T>(table.Rows);

			for(int row = 0; row<table.Rows; row++)
			{
				T rowObj = new T();

				ttm.OnNewRow(rowObj);

				for( int col = 0; col < table.Columns; col++ )
				{
					string value = table.GetValue(row, col);
					if( value != null )
					{
						bool ret = ttm.SetValue(rowObj, col, value);

						// if the value is not handled, give oppotunity for a custom parser
						if( ret == false && context.customParser != null )
						{
							context.customParser(rowObj, table.GetColumnName(col), value);
						}
					}
				}

				result.Add(rowObj);

				if( context.readMaxRows > 0 && result.Count >= context.readMaxRows )
					break;
			}

			// set context result
			context.table = table;
			context.ttm = ttm;

			return result;
		}

		// public methods

		/// <summary>
		/// Read table into list
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="tableName"></param>
		/// <param name="customParser"></param>
		/// <returns>table list. null if not found</returns>
		public List<T> ReadList<T>(string tableName, CustomParser<T> customParser = null) where T : new()
		{
			var context = new ReadContext<T>()
			{
				customParser = customParser
			};

			return ReadListInternal<T>(tableName, context);
		}

		/// <summary>
		/// Read table into list and return first row
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="tableName"></param>
		/// <param name="customParser"></param>
		/// <returns>first row object. default value if not found.</returns>
		public T ReadSingle<T>(string tableName, CustomParser<T> customParser = null) where T : new()
		{
			var context = new ReadContext<T>()
			{
				customParser = customParser,
				readMaxRows = 1
			};

			var list = ReadListInternal<T>(tableName, context);
			if( list != null && list.Count >= 1 )
			{
				return list[0];
			}
			else
			{
				return default(T);
			}
		}

		public Dictionary<TKey, T> ReadDictionary<TKey, T>(string dataName, CustomParser<T> customParser = null) where T : new()
		{
			return ReadDictionary<TKey, T>(dataName, (string)null, customParser);
		}

		/// <summary>
		/// Read table into dictionary
		/// </summary>
		/// <typeparam name="TKey"></typeparam>
		/// <typeparam name="T"></typeparam>
		/// <param name="tableName">table name</param>
		/// <param name="keyName">column name for key</param>
		/// <param name="customParser">custom parser if any</param>
		/// <returns>table dictionary. null if not found.</returns>
		public Dictionary<TKey, T> ReadDictionary<TKey, T>(string tableName, string keyName, CustomParser<T> customParser = null) where T : new()
		{
			var context = new ReadContext<T>()
			{
				customParser = customParser
			};

			List<T> list = ReadListInternal<T>(tableName, context);
			if( list == null )
				return null;

			var dic = new Dictionary<TKey, T>(list.Count);

			int keyColumn = keyName != null ? context.table.FindColumnIndex(keyName) : 0;
			if( keyColumn == -1 )
				throw new ArgumentException();

			for( int row = 0; row < list.Count; row++ )
			{
				string keyStr = context.table.GetValue(row, keyColumn);

				TKey key = Util.ConvertType<TKey>(keyStr);
				dic[key] = list[row];
			}

			return dic;
		}

		/// <summary>
		/// Read table into dictionary
		/// </summary>
		/// <typeparam name="TKey"></typeparam>
		/// <typeparam name="T"></typeparam>
		/// <param name="tableName">table name</param>
		/// <param name="keySelector">select key value from row object</param>
		/// <param name="customParser">custom parser if any</param>
		/// <returns></returns>
		public Dictionary<TKey, T> ReadDictionary<TKey, T>(string tableName, Func<T, TKey> keySelector, CustomParser<T> customParser = null) where T : new()
		{
			var context = new ReadContext<T>()
			{
				customParser = customParser
			};

			List<T> list = ReadListInternal<T>(tableName, context);
			if( list == null )
				return null;

			var dic = new Dictionary<TKey, T>(list.Count);

			for( int row = 0; row < list.Count; row++ )
			{
				TKey key = keySelector(list[row]);

				dic[key] = list[row];
			}

			return dic;
		}

		/// <summary>
		/// Read just a value
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="tableName">table name</param>
		/// <param name="columnName">column name</param>
		/// <param name="defaultValue">a return value for abnormal case</param>
		/// <returns></returns>
		public T ReadValue<T>(string tableName, string columnName, T defaultValue = default(T))
		{
			// find table
			Table table = FindTable(tableName);
			if( table != null )
			{
				int column = table.FindColumnIndex(columnName);
				if( column != -1 )
				{
					string value = table.GetValue(0, column);
					return Util.ConvertType<T>(value);
				}
			}

			return defaultValue;
		}
	}
}
