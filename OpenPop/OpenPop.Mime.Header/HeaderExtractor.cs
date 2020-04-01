using OpenPop.Common;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Text;

namespace OpenPop.Mime.Header
{
	internal static class HeaderExtractor
	{
		private static int FindHeaderEndPosition(byte[] messageContent)
		{
			if (messageContent == null)
			{
				throw new ArgumentNullException("messageContent");
			}
			using (Stream stream = new MemoryStream(messageContent))
			{
				string value;
				do
				{
					value = StreamUtility.ReadLineAsAscii(stream);
				}
				while (!string.IsNullOrEmpty(value));
				return (int)stream.Position;
			}
		}

		public static void ExtractHeadersAndBody(byte[] fullRawMessage, out MessageHeader headers, out byte[] body, IParsingErrorHandler parsingErrorHandler = null)
		{
			if (fullRawMessage == null)
			{
				throw new ArgumentNullException("fullRawMessage");
			}
			int num = FindHeaderEndPosition(fullRawMessage);
			string @string = Encoding.ASCII.GetString(fullRawMessage, 0, num);
			NameValueCollection headers2 = ExtractHeaders(@string);
			headers = new MessageHeader(headers2, parsingErrorHandler);
			body = new byte[fullRawMessage.Length - num];
			Array.Copy(fullRawMessage, num, body, 0, body.Length);
		}

		public static NameValueCollection ExtractHeaders(string messageContent)
		{
			if (messageContent == null)
			{
				throw new ArgumentNullException("messageContent");
			}
			NameValueCollection nameValueCollection = new NameValueCollection();
			using (StringReader stringReader = new StringReader(messageContent))
			{
				string rawHeader;
				while (!string.IsNullOrEmpty(rawHeader = stringReader.ReadLine()))
				{
					KeyValuePair<string, string> keyValuePair = SeparateHeaderNameAndValue(rawHeader);
					string key = keyValuePair.Key;
					StringBuilder stringBuilder = new StringBuilder(keyValuePair.Value);
					while (IsMoreLinesInHeaderValue(stringReader))
					{
						string text = stringReader.ReadLine();
						if (text == null)
						{
							throw new ArgumentException("This will never happen");
						}
						stringBuilder.Append(text);
					}
					nameValueCollection.Add(key, stringBuilder.ToString());
				}
				return nameValueCollection;
			}
		}

		private static bool IsMoreLinesInHeaderValue(TextReader reader)
		{
			int num = reader.Peek();
			if (num == -1)
			{
				return false;
			}
			char c = (char)num;
			if (c != ' ')
			{
				return c == '\t';
			}
			return true;
		}

		internal static KeyValuePair<string, string> SeparateHeaderNameAndValue(string rawHeader)
		{
			if (rawHeader == null)
			{
				throw new ArgumentNullException("rawHeader");
			}
			string key = string.Empty;
			string value = string.Empty;
			int num = rawHeader.IndexOf(':');
			if (num >= 0 && rawHeader.Length >= num + 1)
			{
				key = rawHeader.Substring(0, num).Trim();
				value = rawHeader.Substring(num + 1).Trim();
			}
			return new KeyValuePair<string, string>(key, value);
		}
	}
}
