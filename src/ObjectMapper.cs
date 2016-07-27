using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Reflection;

namespace ExcelToObject
{
	class ObjectMapper
	{
		ExcelReader mReader;

		public ObjectMapper(ExcelReader reader)
		{
			mReader = reader;
		}

		public bool MapInto(object destObj)
		{
			bool hasError = false;

			foreach( var fi in destObj.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public) )
			{
				var attrs = fi.GetCustomAttributes(typeof(ExcelToObjectAttribute), false);
				ExcelToObjectAttribute attr = attrs.Length > 0 ? (ExcelToObjectAttribute)attrs[0] : null;

				if( attr != null && attr.Ignore == true )
					continue;

				try
				{
					object fieldValue = ReadField(fi, attr);
					fi.SetValue(destObj, fieldValue);
				}
				catch( Exception )
				{
					hasError = true;
				}
			}

			return !hasError;
		}

		object ReadField(FieldInfo fi, ExcelToObjectAttribute attr)
		{
			string tablePath = fi.Name;
			if( attr != null && attr.TablePath != null )
			{
				tablePath = attr.TablePath;
			}

			Type fieldType = fi.FieldType;
			if( fieldType.Name == "List`1" )
			{
				Type itemType = fieldType.GetGenericArguments()[0];

				return ReadList(fieldType, tablePath, itemType, attr);
			}
			else if( fieldType.Name == "Dictionary`2" )
			{
				var subTypes = fieldType.GetGenericArguments();

				Type keyType = subTypes[0];
				Type valueType = subTypes[1];

				return ReadDictionary(fieldType, tablePath, valueType, attr);
			}
			else
			{
				return ReadSingle(fieldType, tablePath, attr);
			}
		}

		object ReadList(Type listType, string tablePath, Type itemType, ExcelToObjectAttribute attr)
		{
			object listObj = Util.New(listType);

			var context = new ExcelReader.ReadContext();

			bool retn = mReader.ReadListInternal(tablePath, itemType, context, listObj as IList);

			return retn ? listObj : null;
		}

		object ReadSingle(Type itemType, string tablePath, ExcelToObjectAttribute attr)
		{
			List<object> listObj = new List<object>();

			var context = new ExcelReader.ReadContext()
			{
				readMaxRows = 1
			};

			bool retn = mReader.ReadListInternal(tablePath, itemType, context, listObj);
			return retn ? listObj[0] : null;
		}

		object ReadDictionary(Type dicType, string tablePath, Type valueType, ExcelToObjectAttribute attr)
		{
			object dicObj = Util.New(dicType);

			var context = new ExcelReader.ReadContext();

			ExcelReader.KeySelector keySelector = null;
			if( attr != null && attr.DictionaryKeyName.IsValid() )
				keySelector = ExcelReader.KeySelector.From(valueType, attr.DictionaryKeyName);

			bool retn = mReader.ReadDictionaryInternal(tablePath, valueType, context, dicObj as IDictionary, keySelector);

			return retn ? dicObj : null;
		}
	}
}
