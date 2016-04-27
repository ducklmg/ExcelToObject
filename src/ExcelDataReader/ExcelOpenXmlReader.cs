using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Globalization;

namespace ExcelDataReader
{
	internal class ExcelOpenXmlReader : IDisposable
	{
		public static List<ExcelToObject.SheetData> ReadSheets(byte[] fileContent)
		{
			using( var reader = new ExcelOpenXmlReader() )
			{
				reader.Initialize(fileContent);

				return reader.GetSheets();
			}
		}

		private const string FILE_sharedStrings = "xl/sharedStrings.xml";
		private const string FILE_styles = "xl/styles.xml";
		private const string FILE_workbook = "xl/workbook.xml";
		private const string FILE_rels = "xl/_rels/workbook.xml.rels";
		private const string XL = "xl/";

		private XlsxWorkbook _workbook;
		private int _depth;
		private int _emptyRowCount;

		private ExcelToObject.ZipExtractor mZipExtractor;

		private XmlReader _xmlReader;
		private Stream _sheetStream;
		private object[] _cellsValues;
		private object[] _savedCellsValues;

		private List<int> _defaultDateTimeStyles;
		private string _namespaceUri;

		internal ExcelOpenXmlReader()
		{
			_defaultDateTimeStyles = new List<int>(new int[]
			{
				14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
			});

		}

		private void ReadGlobals()
		{
			_workbook = new XlsxWorkbook(
				mZipExtractor.GetStream(FILE_workbook),
				mZipExtractor.GetStream(FILE_rels),
				mZipExtractor.GetStream(FILE_sharedStrings),
				mZipExtractor.GetStream(FILE_styles));

			CheckDateTimeNumFmts(_workbook.Styles.NumFmts);
		}

		private void CheckDateTimeNumFmts(List<XlsxNumFmt> list)
		{
			if( list.Count == 0 ) return;

			foreach( XlsxNumFmt numFmt in list )
			{
				if( string.IsNullOrEmpty(numFmt.FormatCode) ) continue;
				string fc = numFmt.FormatCode.ToLower();

				int pos;
				while( (pos = fc.IndexOf('"')) > 0 )
				{
					int endPos = fc.IndexOf('"', pos + 1);

					if( endPos > 0 ) fc = fc.Remove(pos, endPos - pos + 1);
				}

				//it should only detect it as a date if it contains
				//dd mm mmm yy yyyy
				//h hh ss
				//AM PM
				//and only if these appear as "words" so either contained in [ ]
				//or delimted in someway
				//updated to not detect as date if format contains a #
				var formatReader = new FormatReader() { FormatString = fc };
				if( formatReader.IsDateFormatString() )
				{
					_defaultDateTimeStyles.Add(numFmt.Id);
				}
			}
		}

		private void ReadSheetGlobals(XlsxWorksheet sheet)
		{
			if( _xmlReader != null ) _xmlReader.Close();
			if( _sheetStream != null ) _sheetStream.Close();

			string sheetPath = sheet.Path;
			if( sheetPath.StartsWith("/xl/") )
				sheetPath = sheetPath.Substring(4);

			_sheetStream = mZipExtractor.GetStream(XL + sheetPath);

			if( null == _sheetStream ) return;

			_xmlReader = XmlReader.Create(_sheetStream);

			//count rows and cols in case there is no dimension elements
			int rows = 0;
			int cols = 0;

			_namespaceUri = null;
			int biggestColumn = 0; //used when no col elements and no dimension
			while( _xmlReader.Read() )
			{
				if( _xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_worksheet )
				{
					//grab the namespaceuri from the worksheet element
					_namespaceUri = _xmlReader.NamespaceURI;
				}

				if( _xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_dimension )
				{
					string dimValue = _xmlReader.GetAttribute(XlsxWorksheet.A_ref);

					sheet.Dimension = new XlsxDimension(dimValue);
					break;
				}

				//removed: Do not use col to work out number of columns as this is really for defining formatting, so may not contain all columns
				//if (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_col)
				//    cols++;

				if( _xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_row )
					rows++;

				//check cells so we can find size of sheet if can't work it out from dimension or col elements (dimension should have been set before the cells if it was available)
				//ditto for cols
				if( sheet.Dimension == null && cols == 0 && _xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_c )
				{
					var refAttribute = _xmlReader.GetAttribute(XlsxWorksheet.A_r);

					if( refAttribute != null )
					{
						var thisRef = ReferenceHelper.ReferenceToColumnAndRow(refAttribute);
						if( thisRef[1] > biggestColumn )
							biggestColumn = thisRef[1];
					}
				}

			}


			//if we didn't get a dimension element then use the calculated rows/cols to create it
			if( sheet.Dimension == null )
			{
				if( cols == 0 )
					cols = biggestColumn;

				if( rows == 0 || cols == 0 )
				{
					sheet.IsEmpty = true;
					return;
				}

				sheet.Dimension = new XlsxDimension(rows, cols);

				//we need to reset our position to sheet data
				_xmlReader.Close();
				_sheetStream.Close();
				_sheetStream = mZipExtractor.GetStream(XL + sheetPath);
				_xmlReader = XmlReader.Create(_sheetStream);

			}

			//read up to the sheetData element. if this element is empty then there aren't any rows and we need to null out dimension

			_xmlReader.ReadToFollowing(XlsxWorksheet.N_sheetData, _namespaceUri);
			if( _xmlReader.IsEmptyElement )
			{
				sheet.IsEmpty = true;
			}
		}

		private bool ReadSheetRow(XlsxWorksheet sheet)
		{
			if( null == _xmlReader ) return false;

			if( _emptyRowCount != 0 )
			{
				_cellsValues = new object[sheet.ColumnsCount];
				_emptyRowCount--;
				_depth++;

				return true;
			}

			if( _savedCellsValues != null )
			{
				_cellsValues = _savedCellsValues;
				_savedCellsValues = null;
				_depth++;

				return true;
			}

			if( (_xmlReader.NodeType == XmlNodeType.Element && _xmlReader.LocalName == XlsxWorksheet.N_row) ||
				_xmlReader.ReadToFollowing(XlsxWorksheet.N_row, _namespaceUri) )
			{
				_cellsValues = new object[sheet.ColumnsCount];

				int rowIndex = int.Parse(_xmlReader.GetAttribute(XlsxWorksheet.A_r));
				if( rowIndex != (_depth + 1) )
				{
					_emptyRowCount = rowIndex - _depth - 1;
				}
				bool hasValue = false;
				string a_s = String.Empty;
				string a_t = String.Empty;
				string a_r = String.Empty;
				int col = 0;
				int row = 0;

				while( _xmlReader.Read() )
				{
					if( _xmlReader.Depth == 2 ) break;

					if( _xmlReader.NodeType == XmlNodeType.Element )
					{
						hasValue = false;

						if( _xmlReader.LocalName == XlsxWorksheet.N_c )
						{
							a_s = _xmlReader.GetAttribute(XlsxWorksheet.A_s);
							a_t = _xmlReader.GetAttribute(XlsxWorksheet.A_t);
							a_r = _xmlReader.GetAttribute(XlsxWorksheet.A_r);
							XlsxDimension.XlsxDim(a_r, out col, out row);
						}
						else if( _xmlReader.LocalName == XlsxWorksheet.N_v || _xmlReader.LocalName == XlsxWorksheet.N_t )
						{
							hasValue = true;
						}
					}

					if( _xmlReader.NodeType == XmlNodeType.Text && hasValue )
					{
						double number;
						object o = _xmlReader.Value;

						var style = NumberStyles.Any;
						var culture = CultureInfo.InvariantCulture;

						if( double.TryParse(o.ToString(), style, culture, out number) )
							o = number;

						if( null != a_t && a_t == XlsxWorksheet.A_s ) //if string
						{
							o = Helpers.ConvertEscapeChars(_workbook.SST[int.Parse(o.ToString())]);
						} // Requested change 4: missing (it appears that if should be else if)
						else if( null != a_t && a_t == XlsxWorksheet.N_inlineStr ) //if string inline
						{
							o = Helpers.ConvertEscapeChars(o.ToString());
						}
						else if( a_t == "b" ) //boolean
						{
							o = _xmlReader.Value == "1";
						}
						else if( null != a_s ) //if something else
						{
							XlsxXf xf = _workbook.Styles.CellXfs[int.Parse(a_s)];
							if( o != null && o.ToString() != string.Empty && IsDateTimeStyle(xf.NumFmtId) )
								o = Helpers.ConvertFromOATime(number);
							else if( xf.NumFmtId == 49 )
								o = o.ToString();
						}



						if( col - 1 < _cellsValues.Length )
							_cellsValues[col - 1] = o;
					}
				}

				if( _emptyRowCount > 0 )
				{
					_savedCellsValues = _cellsValues;
					return ReadSheetRow(sheet);
				}
				_depth++;

				return true;
			}

			_xmlReader.Close();
			if( _sheetStream != null ) _sheetStream.Close();

			return false;
		}

		private bool IsDateTimeStyle(int styleId)
		{
			return _defaultDateTimeStyles.Contains(styleId);
		}


		public void Initialize(byte[] zipArchive)
		{
			mZipExtractor = new ExcelToObject.ZipExtractor(zipArchive);

			ReadGlobals();
		}

		public List<ExcelToObject.SheetData> GetSheets()
		{
			var dataset = new List<ExcelToObject.SheetData>();

			for( int ind = 0; ind < _workbook.Sheets.Count; ind++ )
			{
				List<object[]> rows = new List<object[]>();

				ReadSheetGlobals(_workbook.Sheets[ind]);

				if( _workbook.Sheets[ind].Dimension == null ) continue;

				_depth = 0;
				_emptyRowCount = 0;

				while( ReadSheetRow(_workbook.Sheets[ind]) )
				{
					rows.Add(_cellsValues);
				}

				string[,] table = new string[rows.Count, _cellsValues.Length];

				for( int r = 0; r < rows.Count; r++ )
				{
					var row = rows[r];
					for( int c = 0; c < row.Length; c++ )
					{
						if( row[c] != null )
						{
							table[r, c] = row[c].ToString();
							if( String.IsNullOrEmpty(table[r, c]) )     // 값이 없으면 null을 리턴하게 하자. Empty스트링도 모두 null로 바꾼다.
								table[r, c] = null;
						}
					}
				}

				dataset.Add(new ExcelToObject.SheetData(_workbook.Sheets[ind].Name, table));
			}

			return dataset;
		}

		public void Dispose()
		{
			if( _xmlReader != null ) _xmlReader.Close();

			if( _sheetStream != null ) _sheetStream.Close();
		}
	}
}