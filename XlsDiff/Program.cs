using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelToObject;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

namespace XlsDiff
{
	static class Program
	{
		[STAThread]
		static void Main(string[] args)
		{
			// Diff
			if( args.Length == 2 )
			{
				string left = args[0];
				string right = args[1];

				var leftRt = new ReadableText(left);
				var rightRt = new ReadableText(right);

				// tortoise-merge
				string arg = "";
				arg += String.Format("/base:\"{0}\" ", leftRt.SaveToTemp());
				arg += String.Format("/basename:\"{0}\" ", Path.GetFileName(left));
				arg += String.Format("/mine:\"{0}\" ", rightRt.SaveToTemp());
				arg += String.Format("/minename:\"{0}\" ", Path.GetFileName(right));

				Process.Start(@"C:\Program Files\TortoiseSVN\bin\TortoiseMerge.exe", arg);
			}

			// Merge
			else if( args.Length == 4 )
			{
				string merged = args[0];
				string @base = args[1];
				string mine = args[2];
				string theirs = args[3];

				var baseRt = new ReadableText(@base);
				var mineRt = new ReadableText(mine);
				var theirsRt = new ReadableText(theirs);

				// tortoise-merge
				string arg = "";
				arg += String.Format("/merged:\"{0}\" ", merged);
				arg += String.Format("/mergedname:\"{0}\" ", Path.GetFileName(merged));
				arg += String.Format("/base:\"{0}\" ", baseRt.SaveToTemp());
				arg += String.Format("/basename:\"{0}\" ", Path.GetFileName(@base));
				arg += String.Format("/mine:\"{0}\" ", mineRt.SaveToTemp());
				arg += String.Format("/minename:\"{0}\" ", Path.GetFileName(mine));
				arg += String.Format("/theirs:\"{0}\" ", theirsRt.SaveToTemp());
				arg += String.Format("/theirsname:\"{0}\" ", Path.GetFileName(theirs));

				Process.Start(@"C:\Program Files\TortoiseSVN\bin\TortoiseMerge.exe", arg);
			}
			else
			{
				RegisterTortoiseSVN();
			}
		}

		static void RegisterTortoiseSVN()
		{
			string diffArg = Application.ExecutablePath + " %base %mine";
			Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\SOFTWARE\TortoiseSVN\DiffTools", ".xlsx", diffArg);

			string mergeArg = Application.ExecutablePath + " %merged %base %mine %theirs";
			Microsoft.Win32.Registry.SetValue(@"HKEY_CURRENT_USER\SOFTWARE\TortoiseSVN\MergeTools", ".xlsx", mergeArg);

			MessageBox.Show("Registered to TortoiseSVN Diff");
		}
	}

	class ReadableText
	{
		StringBuilder mBuilder = new StringBuilder();

		public ReadableText(string xlsxFilePath)
		{
			MakeText(ExcelReader.Read(xlsxFilePath));
		}

		public override string ToString()
		{
			return mBuilder.ToString();
		}

		void MakeText(List<SheetData> sheets)
		{
			foreach( var sheet in sheets )
			{
				var tables = sheet.GetAllTables();
				foreach( var table in tables )
				{
					AddTable(table);

					mBuilder.AppendLine();
					mBuilder.AppendLine();
				}
			}
		}

		void AddString(int width, string value)
		{
			if( value == null )
				value = "";

			mBuilder.Append(value.PadRight(width));
		}

		void AddTable(Table table)
		{
			int col = table.Columns;
			int row = table.Rows;

			mBuilder.AppendFormat("[{0}]", table.Name);
			mBuilder.AppendLine();

			var maxWidths = new int[col];
			for( int c = 0; c < col; c++ )
			{
				int maxWidth = 4;

				string name = table.GetColumnName(c);
				if( maxWidth < name.Length )
					maxWidth = name.Length;

				for( int r = 0; r < row; r++ )
				{
					string value = table.GetValue(r, c);
					if( value != null && maxWidth < value.Length )
						maxWidth = value.Length;
				}

				maxWidths[c] = maxWidth + 2;
			}

			for( int c = 0; c < col; c++ )
			{
				AddString(maxWidths[c], table.GetColumnName(c));
			}
			mBuilder.AppendLine();

			for( int r = 0; r < row; r++ )
			{
				for( int c = 0; c < col; c++ )
				{
					AddString(maxWidths[c], table.GetValue(r, c));
				}
				mBuilder.AppendLine();
			}
		}

		public string SaveToTemp()
		{
			var path = Path.GetTempFileName();

			File.WriteAllText(path, mBuilder.ToString(), Encoding.UTF8);

			return path;
		}
	}
}
