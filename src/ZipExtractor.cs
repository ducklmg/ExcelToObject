using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Unity.IO.Compression;

namespace ExcelToObject
{
	internal class ZipExtractor
	{
		byte[] mZipArchive;

		Dictionary<string, byte[]> mFiles;

		public ZipExtractor(byte[] zipArchive)
		{
			UnzipAllFiles(zipArchive);
		}

		private void UnzipAllFiles(byte[] zipArchive)
		{
			if( zipArchive.Length < 22 )
				throw new ArgumentException();

			mZipArchive = zipArchive;
			mFiles = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

			int fileCount;
			int cdPos = FindCentralDirectory(out fileCount);

			for( int i = 0; i < fileCount; i++ )
			{
				FileHeader header = ReadFileHeader(ref cdPos);
				if( header.uncompressLen > 0 )
				{
					byte[] result = ExtractFile(header);

					mFiles[header.name] = result;
				}
			}
		}

		int FindCentralDirectory(out int fileCount)
		{
			byte[] bytes = mZipArchive;
			int len = mZipArchive.Length;

			int eocdPos = len - 22;

			for( ; eocdPos >= 0; eocdPos-- )
			{
				uint sig = BitConverter.ToUInt32(bytes, eocdPos);
				if( sig == 0x06054b50 )
				{
					int commentLen = BitConverter.ToInt16(bytes, eocdPos + 20);
					if( eocdPos + commentLen + 22 == len )
					{
						int cdPos = BitConverter.ToInt32(bytes, eocdPos + 16);

						sig = BitConverter.ToUInt32(bytes, cdPos);
						if( sig == 0x02014b50 )
						{
							fileCount = BitConverter.ToInt16(bytes, eocdPos + 8);
							return cdPos;
						}
					}
				}
			}

			throw new ArgumentException("Invalid zip file");
		}

		struct FileHeader
		{
			public string name;
			public int compressLen;
			public int uncompressLen;
			public int offset;
			public bool compressed;
		}

		FileHeader ReadFileHeader(ref int cdPos)
		{
			var bytes = mZipArchive;

			int cdCompressLen = BitConverter.ToInt32(bytes, cdPos + 20);
			int cdUncompressLen = BitConverter.ToInt32(bytes, cdPos + 24);
			int nameLen = BitConverter.ToInt16(bytes, cdPos + 28);
			int extraLen = BitConverter.ToInt16(bytes, cdPos + 30);
			int commentLen = BitConverter.ToInt16(bytes, cdPos + 32);
			string name = Encoding.UTF8.GetString(bytes, cdPos + 46, nameLen);

			int headerPos = BitConverter.ToInt32(bytes, cdPos + 42);

			cdPos += 46 + nameLen + extraLen + commentLen;

			uint sig = BitConverter.ToUInt32(bytes, headerPos);
			if( sig != 0x04034b50 )
				throw new ArgumentException("Invalid zip file (local file header signature error)");

			int flag = BitConverter.ToInt16(bytes, headerPos + 6);
			int compressLen = BitConverter.ToInt32(bytes, headerPos + 18);
			int uncompressLen = BitConverter.ToInt32(bytes, headerPos + 22);

			int name2Len = BitConverter.ToInt16(bytes, headerPos + 26);
			int extra2Len = BitConverter.ToInt16(bytes, headerPos + 28);

			int method = BitConverter.ToInt16(bytes, headerPos + 8);
			if( method != 0 && method != 8 )
				throw new NotSupportedException("Not supported zip compression method");

			bool compressed = method != 0;      //0=STORE, 8=DEFLATE (https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT)

			int dataPos = headerPos + 30 + name2Len + extra2Len;
			
			if( (flag & 0x8) == 0x8 )
			{
				// the CRC-32 and file sizes are not known when the header is written.
				// get from the data descriptor placed at the end of compressed data
				int dataDescPos = dataPos + cdCompressLen;
				int dataDescSig = BitConverter.ToInt32(bytes, dataDescPos);
				if( dataDescSig != 0x08074b50 )
					throw new ArgumentException("Invalid zip file (local file optional data descriptor signature error)");

				compressLen = BitConverter.ToInt32(bytes, dataDescPos + 8);
				uncompressLen = BitConverter.ToInt32(bytes, dataDescPos + 12);
			}

			var header = new FileHeader();
			header.name = name;
			header.compressLen = compressLen;
			header.uncompressLen = uncompressLen;
			header.offset = dataPos;
			header.compressed = compressed;

			return header;
		}

		byte[] ExtractFile(FileHeader header)
		{
			var bytes = mZipArchive;

			byte[] dest = new byte[header.uncompressLen];

			if( header.compressed )
			{
				using( var ms = new MemoryStream(bytes, header.offset, header.compressLen) )
				{
					using( var ds = new DeflateStream(ms, CompressionMode.Decompress, true) )
					{
						int readLen = ds.Read(dest, 0, header.uncompressLen);
					}
				}
			}
			else
			{
				Array.Copy(bytes, header.offset, dest, 0, header.compressLen);
			}

			return dest;
		}

		public byte[] GetBytes(string path)
		{
			byte[] data;
			mFiles.TryGetValue(path, out data);

			return data;
		}

		public Stream GetStream(string path)
		{
			byte[] data = GetBytes(path);
			return data != null ? new MemoryStream(data) : null;
		}
	}
}
