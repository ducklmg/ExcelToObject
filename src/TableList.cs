using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Reflection;

namespace ExcelToObject
{
	public class TableList
	{
		Dictionary<string, Table> mTables;

		public TableList(string filePath) : this(ExcelReader.Read(filePath))
		{
		}

		public TableList(IEnumerable<string> filePathList) : this(ExcelReader.Read(filePathList))
		{
		}

		public TableList(byte[] xlsxFile) : this(ExcelReader.Read(xlsxFile))
		{
		}

		public TableList(IEnumerable<byte[]> xlsxFileList) : this(ExcelReader.Read(xlsxFileList))
		{
		}

		public TableList(SheetData sheet) : this(new List<SheetData>() { sheet })
		{
		}

		public TableList(List<SheetData> sheets)
		{
			var allTables = new List<Table>();

			foreach( var sheet in sheets )
			{
				sheet.FindTables(null, allTables);
			}

			// 이름이 같은 테이블은 머지한다.
			var tableDic = new Dictionary<string, List<Table>>();
			foreach( var table in allTables )
			{
				List<Table> list;
				if( tableDic.TryGetValue(table.Name, out list) == false )
				{
					list = new List<Table>();
					tableDic.Add(table.Name, list);
				}

				list.Add(table);
			}

			mTables = new Dictionary<string, Table>();

			foreach( var item in tableDic )
			{
				mTables[item.Key] = MergeTables(item.Value, null);
			}
		}

		public Table this[string name]
		{
			get
			{
				Table table;
				mTables.TryGetValue(name, out table);
				return table;
			}
		}

		public static Table MergeTables(List<Table> allTables, string joinKey)
		{
			if( allTables.Count == 0 )
				return null;

			if( allTables.Count == 1 )
				return allTables[0];

			string[] columnList = allTables.SelectMany(x => x.GetColumnNames()).Distinct().ToArray();
			int index = 0;
			Dictionary<string, int> columnDic = columnList.ToDictionary(x => x, x => index++);

			if( joinKey.IsEmpty() )
				joinKey = allTables[0].GetColumnName(0);

			var mergedTable = new Dictionary<string, string[]>();

			foreach( var table in allTables )
			{
				var indexMap = table.GetColumnNames().Select(x => columnDic[x]).ToArray();

				int keyCol = table.FindColumnIndex(joinKey);
				if( keyCol == -1 )
					throw new Exception(String.Format("Join key '{0}' does not exist on table '{1}'", joinKey, table.Name));

				for( int row = 0; row < table.Rows; row++ )
				{
					string keyValue = table.GetValue(row, keyCol);

					string[] rowValues;
					if( mergedTable.TryGetValue(keyValue, out rowValues) == false )
					{
						rowValues = new string[columnDic.Count];
						mergedTable.Add(keyValue, rowValues);
					}

					for( int col = 0; col < table.Columns; col++ )
					{
						int mappedIndex = indexMap[col];

						string value = table.GetValue(row, col);
						if( value.IsValid() )
							rowValues[mappedIndex] = value;
					}
				}
			}

			return CreateTable(allTables[0].Name, columnList, mergedTable.Values.ToArray());
		}

		static Table CreateTable(string tableName, string[] columns, string[][] mergedRows)
		{
			int rows = mergedRows.Length;
			int cols = columns.Length;

			var data = new string[rows + 2, cols];

			// name
			data[0, 0] = String.Format("[{0}]", tableName);

			// column
			for( int col = 0; col < cols; col++ )
			{
				data[1, col] = columns[col];
			}

			// data
			for( int row = 0; row < rows; row++ )
			{
				for( int col = 0; col < cols; col++ )
				{
					data[row + 2, col] = mergedRows[row][col];
				}
			}

			return new Table(data, 0, 0);
		}
		
		public void MapInto(object destObj)
		{
			StringBuilder error = null;

			foreach( var fi in destObj.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public) )
			{
				var attrs = fi.GetCustomAttributes(typeof(ExcelToObjectAttribute), false);
				ExcelToObjectAttribute attr = attrs.Length > 0 ? (ExcelToObjectAttribute)attrs[0] : ExcelToObjectAttribute.Default;

				if( attr.Ignore == true )
					continue;

				try
				{
					string tableName = fi.Name;
					if( attr.TableName != null )
					{
						tableName = attr.TableName;
					}

					if( attr.MapInto )
					{
						// recursively map into sub-object
						object subObj = Util.New(fi.FieldType);

						MapInto(subObj);

						fi.SetValue(destObj, subObj);
					}
					else
					{
						var table = this[tableName];
						if( table != null )
						{
							object fieldValue = table.ReadField(fi, attr.DictionaryKeyName);
							fi.SetValue(destObj, fieldValue);
						}
					}
				}
				catch( Exception ex)
				{
					if( error == null )
						error = new StringBuilder();

					error.AppendFormat("[{0}:{1}] ", fi.Name, ex.Message);
				}
			}

			if( error != null )
			{
				throw new Exception(error.ToString());
			}
		}
	}
}
