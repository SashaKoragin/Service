using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OpenPop.Mime.Decode
{
	public static class EncodingFinder
	{
		public delegate Encoding FallbackDecoderDelegate(string characterSet);

		public static FallbackDecoderDelegate FallbackDecoder
		{
			private get;
			set;
		}

		private static Dictionary<string, Encoding> EncodingMap
		{
			get;
			set;
		}

		static EncodingFinder()
		{
			Reset();
		}

		internal static void Reset()
		{
			EncodingMap = new Dictionary<string, Encoding>();
			FallbackDecoder = null;
			AddMapping("utf8", Encoding.UTF8);
			AddMapping("binary", Encoding.ASCII);
		}

		internal static Encoding FindEncoding(string characterSet)
		{
			if (characterSet == null)
			{
				throw new ArgumentNullException("characterSet");
			}
			string text = characterSet.ToUpperInvariant();
			if (EncodingMap.ContainsKey(text))
			{
				return EncodingMap[text];
			}
			try
			{
				if (text.Contains("WINDOWS") || text.Contains("CP"))
				{
					text = text.Replace("CP", "");
					text = text.Replace("WINDOWS", "");
					text = text.Replace("-", "");
					int codepage = int.Parse(text, CultureInfo.InvariantCulture);
					return Encoding.GetEncoding(codepage);
				}
				if (text.Length > 3 && text.StartsWith("ISO") && text[3] >= '0' && text[3] <= '9')
				{
					return Encoding.GetEncoding("iso-" + characterSet.Substring(3));
				}
				return Encoding.GetEncoding(characterSet);
			}
			catch (ArgumentException)
			{
				if (FallbackDecoder == null)
				{
					throw;
				}
				Encoding encoding = FallbackDecoder(characterSet);
				if (encoding == null)
				{
					throw;
				}
				return encoding;
			}
		}

		public static void AddMapping(string characterSet, Encoding encoding)
		{
			if (characterSet == null)
			{
				throw new ArgumentNullException("characterSet");
			}
			if (encoding == null)
			{
				throw new ArgumentNullException("encoding");
			}
			EncodingMap.Add(characterSet.ToUpperInvariant(), encoding);
		}
	}
}
