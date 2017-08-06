using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.ComponentModel;

namespace ExcelToObject
{
	static class Util
	{
		public static T ConvertType<T>(string value)
		{
			if( typeof(T).IsEnum )
				return (T)Enum.Parse(typeof(T), value);

			// contribution from gasbank
			var converter = TypeDescriptor.GetConverter(typeof(T));
			if( converter != null && converter.CanConvertFrom(value.GetType()) )
			{
				return (T)converter.ConvertFrom(value);
			}
			else
			{
				return (T)Convert.ChangeType(value, typeof(T));
			}
		}

		public static object ConvertType(string value, Type type)
		{
			if( type.IsEnum )
				return Enum.Parse(type, value);

			// contribution from gasbank
			var converter = TypeDescriptor.GetConverter(type);
			if( converter != null && converter.CanConvertFrom(value.GetType()) )
			{
				return converter.ConvertFrom(value);
			}
			else
			{
				return Convert.ChangeType(value, type);
			}
		}

		public static bool IsValid(this string s)
		{
			return !s.IsEmpty();
		}

		public static bool IsEmpty(this string s)
		{
			return String.IsNullOrEmpty(s);
		}

		public static byte[] ReadAll(this Stream stream)
		{
			byte[] buffer = new byte[16 * 1024];
			using( MemoryStream ms = new MemoryStream() )
			{
				int read;
				while( (read = stream.Read(buffer, 0, buffer.Length)) != 0 )
					ms.Write(buffer, 0, read);

				return ms.ToArray();
			}
		}

		public static object New(Type type, params object[] args)
		{
			return Activator.CreateInstance(type, args);
		}
	}
}
