using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Collections;

namespace ExcelToObject
{
	/* Table structure in cell grid
	 * 
	 * |[Name]   |
	 * |column1  |column2  |column3  | ...  |
	 * |value1   |value2   |value3   | ...  |
	 * | ...     | ...     | ...     | ...  |
	 * 
	*/
	public class Table
	{

		string[,] mData;
		string mName;

		int mPivotRow;
		int mPivotCol;
		int mColumnCount;
		int mRowCount;

		bool mTransposed;

		public string Name { get { return mName; } }
		public int Columns { get { return mColumnCount; } }
		public int Rows { get { return mRowCount; } }       // rows count (except table name/column row)

		public Table(string[,] data, int pivotRow, int pivotCol)
		{
			mData = data;
			mPivotRow = pivotRow;
			mPivotCol = pivotCol;

			if( ParseTableMarker(Get(0, 0), out mName, out mTransposed) )
			{
				CalcTableSize();
			}
		}

		string Get(int row, int col)
		{
			return mData[mPivotRow + row, mPivotCol + col];
		}

		public static bool IsTableMarker(string value, string match = null)
		{
			string name;
			bool transpose;

			if( ParseTableMarker(value, out name, out transpose) )
			{
				if( match.IsValid() )
					return name == match;

				return true;
			}

			return false;
		}

		static bool ParseTableMarker(string marker, out string name, out bool transpose)
		{
			name = null;
			transpose = false;

			if( marker.IsValid() )
			{
				if( marker[0] == '[' )
				{
					int last = marker.Length - 1;

					if( marker[last] == '*' )
					{
						transpose = true;
						last--;
					}

					if( marker[last] == ']' )
					{
						name = marker.Substring(1, last - 1);
						return true;
					}
				}
			}

			return false;
		}

		void CalcTableSize()
		{
			int maxRows = mData.GetLength(0);
			int maxCols = mData.GetLength(1);

			int startCol = mPivotCol;
			int lastCol;
			for( lastCol = startCol; lastCol < maxCols; lastCol++ )
			{
				if( mData[mPivotRow + 1, lastCol].IsEmpty() )
					break;
			}

			int startRow = mPivotRow + 1;
			int lastRow;
			for( lastRow = startRow; lastRow < maxRows; lastRow++ )
			{
				bool emptyRow = true;
				for( int col = mPivotCol; col < lastCol; col++ )
				{
					if( mData[lastRow, col].IsValid() )
					{
						emptyRow = false;
						break;
					}
				}

				if( emptyRow )
					break;
			}

			if( mTransposed )
			{
				mColumnCount = lastRow - startRow;
				mRowCount = Math.Max(0, lastCol - startCol - 1);
			}
			else
			{
				mColumnCount = lastCol - startCol;
				mRowCount = Math.Max(0, lastRow - startRow - 1);
			}
		}

		public string GetColumnName(int index)
		{
			if( mTransposed )
				return Get(index + 1, 0);

			return Get(1, index);
		}

		public string[] GetColumnNames()
		{
			var names = new string[mColumnCount];

			for( int i = 0; i < mColumnCount; i++ )
			{
				names[i] = GetColumnName(i);
			}

			return names;
		}

		public int FindColumnIndex(string name)
		{
			for( int i = 0; i < mColumnCount; i++ )
			{
				string n = GetColumnName(i);
				if( n == name )
					return i;
			}

			return -1;
		}

		public string GetValue(int row, int col)
		{
			if( mTransposed )
				return Get(col + 1, row + 1);

			return Get(row + 2, col);
		}

		//////////////////////////////////////////////////////////////
		internal bool ReadListInternal(Type itemType, int readMaxRows, IList destList)
		{
			// header parse
			var ttm = new TableToTypeMap(itemType);

			for( int col = 0; col < mColumnCount; col++ )
			{
				ttm.AddFieldColumn(GetColumnName(col));
			}

			int readRows = mRowCount;
			if( 0 < readMaxRows && readMaxRows < readRows )
				readRows = readMaxRows;

			// rows parse
			for( int row = 0; row < readRows; row++ )
			{
				object rowObj = Util.New(itemType);

				ttm.OnNewRow(rowObj);

				for( int col = 0; col < mColumnCount; col++ )
				{
					string value = GetValue(row, col);
					if( value != null )
					{
						bool ret = ttm.SetValue(rowObj, col, value);
					}
				}

				destList.Add(rowObj);
			}

			return true;
		}

		private List<T> ReadListInternal<T>(int readMaxRows) where T : new()
		{
			var result = new List<T>();

			bool retn = ReadListInternal(typeof(T), readMaxRows, result);

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
		public List<T> ReadList<T>() where T : new()
		{
			return ReadListInternal<T>(0);
		}

		/// <summary>
		/// Read table into list and return first row
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="tablePath"></param>
		/// <param name="customParser"></param>
		/// <returns>first row object. default value if not found.</returns>
		public T ReadSingle<T>() where T : new()
		{
			var list = ReadListInternal<T>(1);
			return list != null ? list[0] : default(T);
		}


		#region ReadDictionary

		internal class KeySelector
		{
			Func<object, object> mFunc;
			FieldInfo mFi;

			static public KeySelector From<TKey, TValue>(Func<TValue, TKey> selectFunc)
			{
				return new KeySelector() { mFunc = (value => selectFunc((TValue)value)) };
			}

			static public KeySelector From(Type valueType, string keyName)
			{
				return new KeySelector() { mFi = valueType.GetField(keyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic) };
			}

			public object Select(object value)
			{
				if( mFi != null )
					return mFi.GetValue(value);

				if( mFunc != null )
					return mFunc(value);

				throw new Exception();
			}
		}

		internal bool ReadDictionaryInternal(Type valueType, IDictionary destDic, KeySelector keySelector = null)
		{
			List<object> values = new List<object>();

			bool retn = ReadListInternal(valueType, 0, values);
			if( retn == false )
				return false;

			if( keySelector == null )
				keySelector = KeySelector.From(valueType, GetColumnName(0));

			foreach( object value in values )
			{
				object key = keySelector.Select(value);

				try
				{
					destDic.Add(key, value);
				}
				catch( ArgumentException ae )
				{
					// Key already exists?
					throw new Exception(string.Format("Duplicated key : tablePath={0}, key={1}", mName, key), ae);
				}
			}

			return true;
		}

		public Dictionary<TKey, T> ReadDictionary<TKey, T>(string keyName) where T : new()
		{
			var dic = new Dictionary<TKey, T>();

			var keySelector = keyName.IsValid() ? KeySelector.From(typeof(T), keyName) : null;

			bool retn = ReadDictionaryInternal(typeof(T), dic, keySelector);

			return retn ? dic : null;
		}

		public Dictionary<TKey, T> ReadDictionary<TKey, T>() where T : new()
		{
			return ReadDictionary<TKey, T>((string)null);
		}

		public Dictionary<TKey, T> ReadDictionary<TKey, T>(Func<T, TKey> keySelector) where T : new()
		{
			var dic = new Dictionary<TKey, T>();

			bool retn = ReadDictionaryInternal(typeof(T), dic, KeySelector.From(keySelector));

			return retn ? dic : null;
		}

		#endregion

		public object ReadField(FieldInfo fi, string dicKeyName = null)
		{
			Type fieldType = fi.FieldType;
			if( fieldType.Name == "List`1" )
			{
				Type itemType = fieldType.GetGenericArguments()[0];

				object listObj = Util.New(fieldType);

				bool retn = ReadListInternal(itemType, 0, listObj as IList);

				return retn ? listObj : null;
			}
			else if( fieldType.Name == "Dictionary`2" )
			{
				var subTypes = fieldType.GetGenericArguments();

				Type keyType = subTypes[0];
				Type valueType = subTypes[1];

				object dicObj = Util.New(fieldType);

				KeySelector keySelector = dicKeyName.IsValid() ? KeySelector.From(valueType, dicKeyName) : null;

				bool retn = ReadDictionaryInternal(valueType, dicObj as IDictionary, keySelector);

				return retn ? dicObj : null;
			}
			else
			{
				List<object> listObj = new List<object>();

				bool retn = ReadListInternal(fieldType, 1, listObj);

				return retn ? listObj[0] : null;
			}
		}
	}
}
