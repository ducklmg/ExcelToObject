using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.IO;
using System.Reflection;
using ExcelDataReader;

namespace ExcelToObject
{
	/// <summary>
	/// Read and parse excel table.
	/// </summary>
	public class ExcelReader
	{
		List<SheetData> mSheets;
		public enum ReadMode
		{
			ExclusiveRead,
			SharedRead,
		}

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

		public ExcelReader(string filePath, ReadMode readMode)
		{
			if (readMode == ReadMode.SharedRead)
			{
				using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
				{
					ReadSheetList(fs.ReadAll());
				}
			}
			else
			{
				ReadSheetList(File.ReadAllBytes(filePath));
			}
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
		/// <param name="path">Table path</param>
		/// <returns>Table instance. If not found, returns null.</returns>
		public Table FindTable(string path)
		{
			/* path syntax
			 *  1) TableName : find all sheets for the name.
			 *  2) SheetName/TableName : find just in the sheet.
			 */
			int pos = path.IndexOf('/');

			string sheetName = pos != -1 ? path.Substring(0, pos) : null;
			string tableName = pos != -1 ? path.Substring(pos+1) : path;

			for( int i = 0; i < mSheets.Count; i++ )
			{
				if( sheetName.IsValid() && mSheets[i].Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == false )
					continue;

				var table = mSheets[i].FindTable(tableName);
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
		public delegate bool CustomParser(object obj, string name, string value);

		//////////////////////////////////////////////////////////////
		internal class ReadContext
		{
			public int readMaxRows;
			public CustomParser customParser;

			public Table table;
			public TableToTypeMap ttm;
		}

		internal bool ReadListInternal(string tablePath, Type itemType, ReadContext context, IList destList)
		{
			// find table
			Table table = FindTable(tablePath);
			if( table == null )
				return false;

			// header parse
			var ttm = new TableToTypeMap(itemType);

			for( int col = 0; col < table.Columns; col++ )
			{
				ttm.AddFieldColumn(table.GetColumnName(col));
			}

			int readRows = table.Rows;
			if( 0 < context.readMaxRows && context.readMaxRows < table.Rows )
				readRows = context.readMaxRows;

			// rows parse
			for( int row = 0; row < readRows; row++ )
			{
				object rowObj = Util.New(itemType);

				ttm.OnNewRow(rowObj);

				for( int col = 0; col < table.Columns; col++ )
				{
					string value = table.GetValue(row, col);
					if( value != null )
					{
						bool ret = ttm.SetValue(rowObj, col, value);

						// if the value is not handled, give oppotunity to a custom parser
						if( ret == false && context.customParser != null )
						{
							context.customParser(rowObj, table.GetColumnName(col), value);
						}
					}
				}

				destList.Add(rowObj);
			}

			// set context result
			context.table = table;
			context.ttm = ttm;

			return true;
		}

		private List<T> ReadListInternal<T>(string tablePath, ReadContext context) where T : new()
		{
			var result = new List<T>();

			bool retn = ReadListInternal(tablePath, typeof(T), context, result);

			result.TrimExcess();
			return retn ? result : null;
		}

		/// <summary>
		/// Read table into list
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="tablePath"></param>
		/// <param name="customParser"></param>
		/// <returns>table list. null if not found</returns>
		public List<T> ReadList<T>(string tablePath, CustomParser customParser = null) where T : new()
		{
			var context = new ReadContext()
			{
				customParser = customParser
			};

			return ReadListInternal<T>(tablePath, context);
		}

		/// <summary>
		/// Read table into list and return first row
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="tablePath"></param>
		/// <param name="customParser"></param>
		/// <returns>first row object. default value if not found.</returns>
		public T ReadSingle<T>(string tablePath, CustomParser customParser = null) where T : new()
		{
			var context = new ReadContext()
			{
				customParser = customParser,
				readMaxRows = 1
			};

			var list = ReadListInternal<T>(tablePath, context);
			return list != null ? list[0] : default(T);
		}



		#region ReadDictionary

		internal class KeySelector
		{
			Func<object, object> func;
			FieldInfo fi;

			static public KeySelector From<TKey, TValue>(Func<TValue, TKey> selectFunc)
			{
				return new KeySelector() { func = (value => selectFunc((TValue)value)) };
			}

			static public KeySelector From(Type valueType, string keyName)
			{
				return new KeySelector() { fi = valueType.GetField(keyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic) };
			}

			public object Select(object value)
			{
				if( fi != null )
					return fi.GetValue(value);

				if( func != null )
					return func(value);

				throw new Exception();
			}
		}

		internal bool ReadDictionaryInternal(string tablePath, Type valueType, ReadContext context, IDictionary destDic, KeySelector keySelector = null)
		{
			List<object> values = new List<object>();

			bool retn = ReadListInternal(tablePath, valueType, context, values);
			if( retn == false )
				return false;

			if( keySelector == null )
				keySelector = KeySelector.From(valueType, context.table.GetColumnName(0));

			foreach( object value in values )
			{
				object key = keySelector.Select(value);

				try
				{
					destDic.Add(key, value);
				}
				catch (ArgumentException ae)
				{
					// Key already exists?
					throw new Exception(string.Format("tablePath={0}, key={1}", tablePath, key), ae);
				}
			}

			return true;
		}

		public Dictionary<TKey, T> ReadDictionary<TKey, T>(string tablePath, string keyName, CustomParser customParser = null) where T : new()
		{
			var context = new ReadContext()
			{
				customParser = customParser
			};

			var dic = new Dictionary<TKey, T>();

			var keySelector = keyName.IsValid() ? KeySelector.From(typeof(T), keyName) : null;

			bool retn = ReadDictionaryInternal(tablePath, typeof(T), context, dic, keySelector);

			return retn ? dic : null;
		}

		public Dictionary<TKey, T> ReadDictionary<TKey, T>(string tablePath, CustomParser customParser = null) where T : new()
		{
			return ReadDictionary<TKey, T>(tablePath, String.Empty, customParser);
		}

		public Dictionary<TKey, T> ReadDictionary<TKey, T>(string tablePath, Func<T, TKey> keySelector, CustomParser customParser = null) where T : new()
		{
			var context = new ReadContext()
			{
				customParser = customParser
			};

			var dic = new Dictionary<TKey, T>();

			bool retn = ReadDictionaryInternal(tablePath, typeof(T), context, dic, KeySelector.From(keySelector));

			return retn ? dic : null;
		}

		#endregion

		/// <summary>
		/// Read just a value
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="tablePath">table name</param>
		/// <param name="columnName">column name</param>
		/// <param name="defaultValue">a return value for abnormal case</param>
		/// <returns></returns>
		public T ReadValue<T>(string tablePath, string columnName, T defaultValue = default(T))
		{
			// find table
			Table table = FindTable(tablePath);
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

		public void MapInto(object destObj)
		{
			var mapper = new ObjectMapper(this);

			bool retn = mapper.MapInto(destObj);

#if DEBUG
			if( retn == false )
				throw new Exception();
#endif
		}
	}
}
