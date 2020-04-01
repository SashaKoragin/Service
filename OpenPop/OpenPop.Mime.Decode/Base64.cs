using OpenPop.Common.Logging;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace OpenPop.Mime.Decode
{
	internal static class Base64
	{
		public static byte[] Decode(string base64Encoded)
		{
			try
			{
				using (MemoryStream memoryStream = new MemoryStream())
				{
					base64Encoded = base64Encoded.Replace("\r\n", "");
					base64Encoded = base64Encoded.Replace("\t", "");
					base64Encoded = base64Encoded.Replace(" ", "");
					byte[] bytes = Encoding.ASCII.GetBytes(base64Encoded);
					using (FromBase64Transform fromBase64Transform = new FromBase64Transform(FromBase64TransformMode.DoNotIgnoreWhiteSpaces))
					{
						byte[] array = new byte[fromBase64Transform.OutputBlockSize];
						int num = 0;
						while (bytes.Length - num > 4)
						{
							fromBase64Transform.TransformBlock(bytes, num, 4, array, 0);
							num += 4;
							memoryStream.Write(array, 0, fromBase64Transform.OutputBlockSize);
						}
						array = fromBase64Transform.TransformFinalBlock(bytes, num, bytes.Length - num);
						memoryStream.Write(array, 0, array.Length);
					}
					return memoryStream.ToArray();
				}
			}
			catch (FormatException ex)
			{
				DefaultLogger.Log.LogError("Base64: (FormatException) " + ex.Message + "\r\nOn string: " + base64Encoded);
				throw;
			}
		}

		public static string Decode(string base64Encoded, Encoding encoding)
		{
			if (base64Encoded == null)
			{
				throw new ArgumentNullException("base64Encoded");
			}
			if (encoding == null)
			{
				throw new ArgumentNullException("encoding");
			}
			return encoding.GetString(Decode(base64Encoded));
		}
	}
}
