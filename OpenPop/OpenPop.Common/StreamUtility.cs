using System;
using System.IO;
using System.Text;

namespace OpenPop.Common
{
	internal static class StreamUtility
	{
		public static byte[] ReadLineAsBytes(Stream stream)
		{
			if (stream == null)
			{
				throw new ArgumentNullException("stream");
			}
			using (MemoryStream memoryStream = new MemoryStream())
			{
				char c;
				do
				{
					int num = stream.ReadByte();
					if (num == -1 && memoryStream.Length > 0)
					{
						break;
					}
					if (num == -1 && memoryStream.Length == 0)
					{
						return null;
					}
					c = (char)num;
					if (c != '\r' && c != '\n')
					{
						memoryStream.WriteByte((byte)num);
					}
				}
				while (c != '\n');
				return memoryStream.ToArray();
			}
		}

		public static string ReadLineAsAscii(Stream stream)
		{
			byte[] array = ReadLineAsBytes(stream);
			if (array == null)
			{
				return null;
			}
			return Encoding.ASCII.GetString(array);
		}
	}
}
