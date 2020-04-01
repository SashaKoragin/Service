using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace OpenPop.Mime.Decode
{
	internal static class QuotedPrintable
	{
		public static string DecodeEncodedWord(string toDecode, Encoding encoding)
		{
			if (toDecode == null)
			{
				throw new ArgumentNullException("toDecode");
			}
			if (encoding == null)
			{
				throw new ArgumentNullException("encoding");
			}
			return encoding.GetString(Rfc2047QuotedPrintableDecode(toDecode, encodedWordVariant: true));
		}

		public static byte[] DecodeContentTransferEncoding(string toDecode)
		{
			if (toDecode == null)
			{
				throw new ArgumentNullException("toDecode");
			}
			return Rfc2047QuotedPrintableDecode(toDecode, encodedWordVariant: false);
		}

		private static byte[] Rfc2047QuotedPrintableDecode(string toDecode, bool encodedWordVariant)
		{
			if (toDecode == null)
			{
				throw new ArgumentNullException("toDecode");
			}
			using (MemoryStream memoryStream = new MemoryStream())
			{
				toDecode = RemoveIllegalControlCharacters(toDecode);
				for (int i = 0; i < toDecode.Length; i++)
				{
					char c = toDecode[i];
					if (c == '=')
					{
						if (toDecode.Length - i < 3)
						{
							WriteAllBytesToStream(memoryStream, DecodeEqualSignNotLongEnough(toDecode.Substring(i)));
							break;
						}
						string decode = toDecode.Substring(i, 3);
						WriteAllBytesToStream(memoryStream, DecodeEqualSign(decode));
						i += 2;
					}
					else if (c == '_' && encodedWordVariant)
					{
						memoryStream.WriteByte(32);
					}
					else
					{
						memoryStream.WriteByte((byte)c);
					}
				}
				return memoryStream.ToArray();
			}
		}

		private static void WriteAllBytesToStream(Stream stream, byte[] toWrite)
		{
			stream.Write(toWrite, 0, toWrite.Length);
		}

		private static string RemoveIllegalControlCharacters(string input)
		{
			if (input == null)
			{
				throw new ArgumentNullException("input");
			}
			input = RemoveCarriageReturnAndNewLinewIfNotInPair(input);
			return Regex.Replace(input, "[\0-\b\v\f\u000e-\u001f\u007f]", "");
		}

		private static string RemoveCarriageReturnAndNewLinewIfNotInPair(string input)
		{
			if (input == null)
			{
				throw new ArgumentNullException("input");
			}
			StringBuilder stringBuilder = new StringBuilder(input.Length);
			for (int i = 0; i < input.Length; i++)
			{
				if ((input[i] != '\r' || (i + 1 < input.Length && input[i + 1] == '\n')) && (input[i] != '\n' || (i - 1 >= 0 && input[i - 1] == '\r')))
				{
					stringBuilder.Append(input[i]);
				}
			}
			return stringBuilder.ToString();
		}

		private static byte[] DecodeEqualSignNotLongEnough(string decode)
		{
			if (decode == null)
			{
				throw new ArgumentNullException("decode");
			}
			if (decode.Length >= 3)
			{
				throw new ArgumentException("decode must have length lower than 3", "decode");
			}
			if (decode.Length <= 0)
			{
				throw new ArgumentException("decode must have length lower at least 1", "decode");
			}
			if (decode[0] != '=')
			{
				throw new ArgumentException("First part of decode must be an equal sign", "decode");
			}
			return Encoding.ASCII.GetBytes(decode);
		}

		private static byte[] DecodeEqualSign(string decode)
		{
			if (decode == null)
			{
				throw new ArgumentNullException("decode");
			}
			if (decode.Length != 3)
			{
				throw new ArgumentException("decode must have length 3", "decode");
			}
			if (decode[0] != '=')
			{
				throw new ArgumentException("decode must start with an equal sign", "decode");
			}
			if (decode.Contains("\r\n"))
			{
				return new byte[0];
			}
			try
			{
				string value = decode.Substring(1);
				return new byte[1]
				{
					Convert.ToByte(value, 16)
				};
			}
			catch (FormatException)
			{
				return Encoding.ASCII.GetBytes(decode);
			}
		}
	}
}
