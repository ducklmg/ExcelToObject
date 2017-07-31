using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Reflection;

namespace ExcelToObject
{
	enum CollectionType
	{
		None,
		Array,
		List,
		Dictionary
	}

	class ColumnData
	{
		public FieldData fieldData;
		public object dicKeyValue;
	}

	class FieldData
	{
		public FieldInfo fieldInfo;
		public CollectionType collectionType;
		public Type valueType;
		public Type dicKeyType;

		public int columnCount;

		// instance data
		public int arrayIndex;
		public Array array;
		public IList list;
		public IDictionary dic;

		public FieldData(FieldInfo fi)
		{
			fieldInfo = fi;

			Type fieldType = fi.FieldType;
			if( fieldType.IsArray )
			{
				collectionType = CollectionType.Array;
				valueType = fieldType.GetElementType();
			}
			else if( fieldType.Name == "List`1" )
			{
				collectionType = CollectionType.List;
				var subTypes = fieldType.GetGenericArguments();
				valueType = subTypes[0];
			}
			else if( fieldType.Name == "Dictionary`2" )
			{
				collectionType = CollectionType.Dictionary;
				var subTypes = fieldType.GetGenericArguments();
				dicKeyType = subTypes[0];
				valueType = subTypes[1];
			}
			else
			{
				valueType = fieldType;
			}
		}

		public void OnNewRow(object obj)
		{
			if( collectionType != CollectionType.None )
			{
				object value = fieldInfo.GetValue(obj);
				if( value == null )
				{
					// 여기서 생성해 주자. Array, List, Dictionary 모두 정수 하나를 인자로 받는 생성자를 호출한다.
					value = Util.New(fieldInfo.FieldType, columnCount);

					fieldInfo.SetValue(obj, value);
				}

				if( collectionType == CollectionType.Array )
				{
					array = (Array)value;
					arrayIndex = 0;
				}
				else if( collectionType == CollectionType.List )
				{
					list = value as IList;
				}
				else if( collectionType == CollectionType.Dictionary )
				{
					dic = value as IDictionary;
				}
			}
		}

		public void SetValue(object obj, string valueStr, ColumnData column)
		{
			try
			{
				object typedValue = Util.ConvertType(valueStr, valueType);

				switch( collectionType )
				{
					case CollectionType.None:
						fieldInfo.SetValue(obj, typedValue);
						break;

					case CollectionType.Array:
						if( arrayIndex < array.Length )
							array.SetValue(typedValue, arrayIndex++);
						break;

					case CollectionType.List:
						list.Add(typedValue);
						break;

					case CollectionType.Dictionary:
						dic.Add(column.dicKeyValue, typedValue);
						break;
				}
			}
			catch( ArgumentException )      // from Enum.Parse
			{
				throw new Exception(String.Format("Value string '{0}' is not member of enum {1}", valueStr, valueType.Name));
			}
			catch( Exception )
			{
				throw new Exception(String.Format("Value string '{0}' is not valid for {1}", valueStr, valueType.Name));
			}
		}
	}

	class TableToTypeMap
	{
		Dictionary<string, FieldData> mNameToField = new Dictionary<string, FieldData>();
		List<ColumnData> mColumns = new List<ColumnData>();

		public int ColumnCount { get { return mColumns.Count; } }
		public List<ColumnData> Columns { get { return mColumns; } }

		public TableToTypeMap(Type t)
		{
			foreach( FieldInfo fi in t.GetFields(BindingFlags.Public | BindingFlags.Instance) )
			{
				mNameToField[fi.Name] = new FieldData(fi);
			}
		}

		public void AddFieldColumn(string name)
		{
			string postfix = null;
			int pos = name.IndexOf('#');
			if( pos != -1 )
			{
				postfix = name.Substring(pos + 1).Trim();
				name = name.Substring(0, pos).Trim();
			}
			else
			{
				name = name.Trim();
			}

			ColumnData columnData = null;

			if( name.IsValid() )
			{
				FieldData fieldData = null;
				if( mNameToField.TryGetValue(name, out fieldData) )
				{
					columnData = new ColumnData();

					columnData.fieldData = fieldData;
					fieldData.columnCount++;

					if( fieldData.dicKeyType != null )
						columnData.dicKeyValue = Util.ConvertType(postfix, fieldData.dicKeyType);
				}
			}

			mColumns.Add(columnData);
		}

		public bool SetValue(object instance, int columnIndex, string value)
		{
			if( columnIndex >= mColumns.Count )
				throw new ArgumentException();

			ColumnData cd = mColumns[columnIndex];
			if( cd != null )
			{
				cd.fieldData.SetValue(instance, value, cd);
				return true;
			}

			return false;
		}

		public void OnNewRow(object obj)
		{
			foreach( var field in mNameToField )
			{
				field.Value.OnNewRow(obj);
			}
		}
	}
}
