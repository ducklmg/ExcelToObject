using System;
using System.Collections.Generic;
using System.Text;
using ExcelToObject;
using System.IO;

namespace Test
{
	class Program
	{
		static void Main(string[] args)
		{
			var data = new DataSet();

			var tableList = new TableList("test.xlsx");

			tableList.MapInto(data);

			string json = Newtonsoft.Json.JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
			Console.WriteLine(json);
		}
	}

	public enum Subject
	{
		Math,
		Science,
		English
	}

	public class SimpleData
	{
		public string id;
		public string name;
		public int age;
	}

	public class ArrayData
	{
		public string id;
		public string[] arr;
	}

	public class DictionaryData
	{
		public string id;
		public Dictionary<Subject, float> score;
	}

	public class AuthorData
	{
		public string company;
		public string author;
		public string job;
	}

	public class JoinData
	{
		public string id;
		public string first;
		public string second;
		public string third;
		public float fourth;
		public int fifth;
	}

	public class DataSet
	{
		public List<SimpleData> SimpleData;

		public List<ArrayData> ArrayData;

		public List<DictionaryData> DictionaryData;

		[ExcelToObject(TableName = "RowColumnData")]
		public AuthorData AuthorData;

		public Dictionary<string, JoinData> Join;
	}
}

