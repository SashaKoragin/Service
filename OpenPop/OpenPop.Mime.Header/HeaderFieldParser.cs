using OpenPop.Common.Logging;
using OpenPop.Mime.Decode;
using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Net.Mime;

namespace OpenPop.Mime.Header
{
	internal static class HeaderFieldParser
	{
		public static ContentTransferEncoding ParseContentTransferEncoding(string headerValue)
		{
			if (headerValue == null)
			{
				throw new ArgumentNullException("headerValue");
			}
			switch (headerValue.Trim().ToUpperInvariant())
			{
			case "7BIT":
				return ContentTransferEncoding.SevenBit;
			case "8BIT":
				return ContentTransferEncoding.EightBit;
			case "QUOTED-PRINTABLE":
				return ContentTransferEncoding.QuotedPrintable;
			case "BASE64":
				return ContentTransferEncoding.Base64;
			case "BINARY":
				return ContentTransferEncoding.Binary;
			default:
				DefaultLogger.Log.LogDebug("Wrong ContentTransferEncoding was used. It was: " + headerValue);
				return ContentTransferEncoding.SevenBit;
			}
		}

		public static MailPriority ParseImportance(string headerValue)
		{
			if (headerValue == null)
			{
				throw new ArgumentNullException("headerValue");
			}
			switch (headerValue.ToUpperInvariant())
			{
			case "5":
			case "HIGH":
				return MailPriority.High;
			case "3":
			case "NORMAL":
				return MailPriority.Normal;
			case "1":
			case "LOW":
				return MailPriority.Low;
			default:
				DefaultLogger.Log.LogDebug("HeaderFieldParser: Unknown importance value: \"" + headerValue + "\". Using default of normal importance.");
				return MailPriority.Normal;
			}
		}

		public static ContentType ParseContentType(string headerValue)
		{
			if (headerValue == null)
			{
				throw new ArgumentNullException("headerValue");
			}
			ContentType contentType = new ContentType();
			List<KeyValuePair<string, string>> list = Rfc2231Decoder.Decode(headerValue);
			bool flag = false;
			foreach (KeyValuePair<string, string> item in list)
			{
				string text = item.Key.ToUpperInvariant().Trim();
				string text2 = Utility.RemoveQuotesIfAny(item.Value.Trim());
				switch (text)
				{
				case "":
					if (!flag)
					{
						string text3 = text2.ToUpperInvariant();
						if (text3.Equals("TEXT") || text3.Equals("TEXT/"))
						{
							text2 = "text/plain";
						}
						contentType.MediaType = cleanMediaType(text2);
						flag = true;
					}
					else
					{
						DefaultLogger.Log.LogDebug("Multiple values without a key in the Content Type header. Only the first will get used as the media type");
					}
					break;
				case "BOUNDARY":
					contentType.Boundary = text2;
					break;
				case "CHARSET":
					contentType.CharSet = text2;
					break;
				case "NAME":
					contentType.Name = EncodedWord.Decode(text2);
					break;
				default:
					if (contentType.Parameters == null)
					{
						throw new Exception("The ContentType parameters property is null. This will never be thrown.");
					}
					contentType.Parameters.Add(text, text2);
					break;
				}
			}
			return contentType;
		}

		public static ContentDisposition ParseContentDisposition(string headerValue)
		{
			if (headerValue == null)
			{
				throw new ArgumentNullException("headerValue");
			}
			ContentDisposition contentDisposition = new ContentDisposition();
			List<KeyValuePair<string, string>> list = Rfc2231Decoder.Decode(headerValue);
			foreach (KeyValuePair<string, string> item in list)
			{
				string text = item.Key.ToUpperInvariant().Trim();
				string text2 = Utility.RemoveQuotesIfAny(item.Value.Trim());
				switch (text)
				{
				case "":
					contentDisposition.DispositionType = text2;
					break;
				case "NAME":
				case "FILENAME":
					contentDisposition.FileName = EncodedWord.Decode(text2);
					break;
				case "CREATION-DATE":
				{
					DateTime dateTime6 = contentDisposition.CreationDate = new DateTime(Rfc2822DateTime.StringToDate(text2).Ticks);
					break;
				}
				case "MODIFICATION-DATE":
				{
					DateTime dateTime4 = contentDisposition.ModificationDate = new DateTime(Rfc2822DateTime.StringToDate(text2).Ticks);
					break;
				}
				case "READ-DATE":
				{
					DateTime dateTime2 = contentDisposition.ReadDate = new DateTime(Rfc2822DateTime.StringToDate(text2).Ticks);
					break;
				}
				case "SIZE":
					contentDisposition.Size = SizeParser.Parse(text2);
					break;
				default:
					if (!text.StartsWith("X-"))
					{
						throw new ArgumentException("Unknown parameter in Content-Disposition. Ask developer to fix! Parameter: " + text);
					}
					contentDisposition.Parameters.Add(text, text2);
					break;
				case "CHARSET":
					break;
				}
			}
			return contentDisposition;
		}

		public static string ParseId(string headerValue)
		{
			return headerValue.Trim().TrimEnd('>').TrimStart('<');
		}

		public static List<string> ParseMultipleIDs(string headerValue)
		{
			List<string> list = new List<string>();
			string[] array = headerValue.Trim().Split(new char[1]
			{
				'>'
			}, StringSplitOptions.RemoveEmptyEntries);
			string[] array2 = array;
			foreach (string headerValue2 in array2)
			{
				list.Add(ParseId(headerValue2));
			}
			return list;
		}

		private static string cleanMediaType(string mediaType)
		{
			if (mediaType == null)
			{
				throw new ArgumentNullException("mediaType");
			}
			int num = mediaType.IndexOf('/');
			if (num == -1)
			{
				throw new ArgumentException("Media Type must be in the format type \"/\" subtype", "mediaType");
			}
			if (num == mediaType.Length - 1)
			{
				throw new ArgumentException("Media Type must contain subtype, which is a madatory field and has no default", "mediaType");
			}
			string text = mediaType.Substring(0, num);
			string s = mediaType.Substring(num + 1, mediaType.Length - text.Length - 1);
			text = stripRfc2045TSpecials(text);
			text = stripRfc822Ctls(text);
			text = text.Replace(" ", "");
			s = stripRfc2045TSpecials(s);
			s = stripRfc822Ctls(s);
			s = s.Replace(" ", "");
			string text2 = text + "/" + s;
			if (text2 != mediaType)
			{
				DefaultLogger.Log.LogDebug("Content-Type Media Type value has been changed since it was invalid. Original value \"" + mediaType + "\". New Value \"" + text2 + "\"");
			}
			return text2;
		}

		private static string stripRfc2045TSpecials(string s)
		{
			if (s == null)
			{
				throw new ArgumentNullException("s");
			}
			string[] array = new string[15]
			{
				"(",
				")",
				"<",
				">",
				"@",
				",",
				";",
				":",
				"\\",
				"\"",
				"/",
				"[",
				"]",
				"?",
				"="
			};
			string[] array2 = array;
			foreach (string oldValue in array2)
			{
				s = s.Replace(oldValue, "");
			}
			return s;
		}

		private static string stripRfc822Ctls(string s)
		{
			if (s == null)
			{
				throw new ArgumentNullException("s");
			}
			for (int i = 0; i <= 31; i++)
			{
				s = s.Replace(((char)i).ToString(), "");
			}
			s = s.Replace('\u007f'.ToString(), "");
			return s;
		}
	}
}
