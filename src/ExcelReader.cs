using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.IO;
using System.Reflection;
using ExcelDataReader;
using System.Linq;

namespace ExcelToObject
{
	/// <summary>
	/// Custom parser signiture for individual properties.
	/// </summary>
	/// <typeparam name="T">Instance type</typeparam>
	/// <param name="obj">instance</param>
	/// <param name="name">property name (column)</param>
	/// <param name="value">property value. null if empty cell.</param>
	/// <returns></returns>
	//public delegate bool CustomParser(object obj, string name, string value);

	/// <summary>
	/// Read and parse excel table.
	/// </summary>
	public class ExcelReader
	{
		// contribution from gasbank
		public enum ReadMode
		{
			SharedRead,
			ExclusiveRead,
		}

		static public List<SheetData> Read(byte[] xlsxFile)
		{
			return ExcelOpenXmlReader.ReadSheets(xlsxFile);
		}

		static public List<SheetData> Read(IEnumerable<byte[]> xlsxFileList)
		{
			var result = new List<SheetData>();
			foreach( var bytes in xlsxFileList )
			{
				AddSheetList(bytes, result);
			}

			return result;
		}

		static public List<SheetData> Read(string filePath, ReadMode readMode = ReadMode.SharedRead)
		{
			return Read(new string[] { filePath }, readMode);
		}

		static public List<SheetData> Read(IEnumerable<string> filePathList, ReadMode readMode = ReadMode.SharedRead)
		{
			var result = new List<SheetData>();

			foreach( string filePath in filePathList )
			{
				// 엑셀에서 열려있을 때는 Share모드로 읽어야만 한다.
				byte[] bytes;

				using( var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, readMode == ReadMode.SharedRead ? FileShare.ReadWrite : FileShare.None) )
				{
					bytes = new byte[(int)fs.Length];
					int read = fs.Read(bytes, 0, bytes.Length);
					if( read != bytes.Length )
						throw new IOException();
				}

				AddSheetList(bytes, result);
			}

			return result;
		}

		static public List<SheetData> Read(Stream stream)
		{
			return Read(stream.ReadAll());
		}

		static void AddSheetList(byte[] xlsxFile, List<SheetData> list)
		{
			list.AddRange(ExcelOpenXmlReader.ReadSheets(xlsxFile));
		}
	}
}