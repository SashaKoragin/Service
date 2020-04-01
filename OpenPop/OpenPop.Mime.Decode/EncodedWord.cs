using System;
using System.Text;
using System.Text.RegularExpressions;

namespace OpenPop.Mime.Decode
{
	internal static class EncodedWord
	{
		public static string Decode(string encodedWords)
		{
			if (encodedWords == null)
			{
				throw new ArgumentNullException("encodedWords");
			}
			encodedWords = Regex.Replace(encodedWords, "(?<first>\\=\\?(?<Charset>\\S+?)\\?(?<Encoding>\\w)\\?(?<Content>.+?)\\?\\=)\\s+(?<second>\\=\\?(?<Charset>\\S+?)\\?(?<Encoding>\\w)\\?(?<Content>.+?)\\?\\=)", "${first}${second}");
			encodedWords = Regex.Replace(encodedWords, "(?<first>\\=\\?(?<Charset>\\S+?)\\?(?<Encoding>\\w)\\?(?<Content>.+?)\\?\\=)\\s+(?<second>\\=\\?(?<Charset>\\S+?)\\?(?<Encoding>\\w)\\?(?<Content>.+?)\\?\\=)", "${first}${second}");
			string text = encodedWords;
			MatchCollection matchCollection = Regex.Matches(encodedWords, "\\=\\?(?<Charset>\\S+?)\\?(?<Encoding>\\w)\\?(?<Content>.+?)\\?\\=");
			foreach (Match item in matchCollection)
			{
				if (item.Success)
				{
					string value = item.Value;
					string value2 = item.Groups["Content"].Value;
					string value3 = item.Groups["Encoding"].Value;
					string value4 = item.Groups["Charset"].Value;
					Encoding encoding = EncodingFinder.FindEncoding(value4);
					string newValue;
					switch (value3.ToUpperInvariant())
					{
					case "B":
						newValue = Base64.Decode(value2, encoding);
						break;
					case "Q":
						newValue = QuotedPrintable.DecodeEncodedWord(value2, encoding);
						break;
					default:
						throw new ArgumentException("The encoding " + value3 + " was not recognized");
					}
					text = text.Replace(value, newValue);
				}
			}
			return text;
		}
	}
}
