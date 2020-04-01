using OpenPop.Common.Logging;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace OpenPop.Mime.Decode
{
	internal static class Rfc2231Decoder
	{
		public static List<KeyValuePair<string, string>> Decode(string toDecode)
		{
			if (toDecode == null)
			{
				throw new ArgumentNullException("toDecode");
			}
			toDecode = Regex.Replace(toDecode, "=\\s*\"(?<value>[^\"]*)\"\\s", "=\"${value}\"; ");
			toDecode = Regex.Replace(toDecode, "^(?<first>[^;\\s]+)\\s(?<second>[^;\\s]+)", "${first}; ${second}");
			List<string> list = Utility.SplitStringWithCharNotInsideQuotes(toDecode.Trim(), ';');
			List<KeyValuePair<string, string>> list2 = new List<KeyValuePair<string, string>>(list.Count);
			foreach (string item in list)
			{
				if (item.Trim().Length != 0)
				{
					string[] array = item.Trim().Split(new char[1]
					{
						'='
					}, 2);
					if (array.Length == 1)
					{
						list2.Add(new KeyValuePair<string, string>("", array[0]));
					}
					else
					{
						if (array.Length != 2)
						{
							throw new ArgumentException("When splitting the part \"" + item + "\" by = there was " + array.Length + " parts. Only 1 and 2 are supported");
						}
						list2.Add(new KeyValuePair<string, string>(array[0], array[1]));
					}
				}
			}
			return DecodePairs(list2);
		}

		private static List<KeyValuePair<string, string>> DecodePairs(List<KeyValuePair<string, string>> pairs)
		{
			if (pairs == null)
			{
				throw new ArgumentNullException("pairs");
			}
			List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>(pairs.Count);
			int count = pairs.Count;
			for (int i = 0; i < count; i++)
			{
				KeyValuePair<string, string> item = pairs[i];
				string key = item.Key;
				string text = Utility.RemoveQuotesIfAny(item.Value);
				if (key.EndsWith("*0", StringComparison.OrdinalIgnoreCase) || key.EndsWith("*0*", StringComparison.OrdinalIgnoreCase))
				{
					string encodingUsed = "notEncoded - Value here is never used";
					if (key.EndsWith("*0*", StringComparison.OrdinalIgnoreCase))
					{
						text = DecodeSingleValue(text, out encodingUsed);
						key = key.Replace("*0*", "");
					}
					else
					{
						key = key.Replace("*0", "");
					}
					StringBuilder stringBuilder = new StringBuilder();
					stringBuilder.Append(text);
					int num = i + 1;
					int num2 = 1;
					while (num < count)
					{
						string key2 = pairs[num].Key;
						string text2 = Utility.RemoveQuotesIfAny(pairs[num].Value);
						if (key2.Equals(key + "*" + num2))
						{
							stringBuilder.Append(text2);
							i++;
						}
						else
						{
							if (!key2.Equals(key + "*" + num2 + "*"))
							{
								break;
							}
							if (encodingUsed != null)
							{
								text2 = DecodeSingleValue(text2, encodingUsed);
							}
							stringBuilder.Append(text2);
							i++;
						}
						num++;
						num2++;
					}
					text = stringBuilder.ToString();
					list.Add(new KeyValuePair<string, string>(key, text));
				}
				else if (key.EndsWith("*", StringComparison.OrdinalIgnoreCase))
				{
					key = key.Replace("*", "");
					text = DecodeSingleValue(text, out string _);
					list.Add(new KeyValuePair<string, string>(key, text));
				}
				else
				{
					list.Add(item);
				}
			}
			return list;
		}

		private static string DecodeSingleValue(string toDecode, out string encodingUsed)
		{
			if (toDecode == null)
			{
				throw new ArgumentNullException("toDecode");
			}
			if (toDecode.IndexOf('\'') == -1)
			{
				DefaultLogger.Log.LogDebug("Rfc2231Decoder: Someone asked me to decode a string which was not encoded - returning raw string. Input: " + toDecode);
				encodingUsed = null;
				return toDecode;
			}
			encodingUsed = toDecode.Substring(0, toDecode.IndexOf('\''));
			toDecode = toDecode.Substring(toDecode.LastIndexOf('\'') + 1);
			return DecodeSingleValue(toDecode, encodingUsed);
		}

		private static string DecodeSingleValue(string valueToDecode, string encoding)
		{
			if (valueToDecode == null)
			{
				throw new ArgumentNullException("valueToDecode");
			}
			if (encoding == null)
			{
				throw new ArgumentNullException("encoding");
			}
			valueToDecode = "=?" + encoding + "?Q?" + valueToDecode.Replace("%", "=") + "?=";
			return EncodedWord.Decode(valueToDecode);
		}
	}
}
