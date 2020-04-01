using OpenPop.Mime.Decode;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OpenPop.Mime.Header
{
	public class Received
	{
		public DateTime Date
		{
			get;
			private set;
		}

		public Dictionary<string, string> Names
		{
			get;
			private set;
		}

		public string Raw
		{
			get;
			private set;
		}

		public Received(string headerValue)
		{
			if (headerValue == null)
			{
				throw new ArgumentNullException("headerValue");
			}
			Raw = headerValue;
			Date = DateTime.MinValue;
			if (headerValue.Contains(";"))
			{
				string inputDate = headerValue.Substring(headerValue.LastIndexOf(";") + 1);
				Date = Rfc2822DateTime.StringToDate(inputDate);
			}
			Names = ParseDictionary(headerValue);
		}

		private static Dictionary<string, string> ParseDictionary(string headerValue)
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string input = headerValue;
			if (headerValue.Contains(";"))
			{
				input = headerValue.Substring(0, headerValue.LastIndexOf(";"));
			}
			input = Regex.Replace(input, "\\s+", " ");
			MatchCollection matchCollection = Regex.Matches(input, "(?<name>[^\\s]+)\\s(?<value>[^\\s]+(\\s\\(.+?\\))*)");
			foreach (Match item in matchCollection)
			{
				string value = item.Groups["name"].Value;
				string value2 = item.Groups["value"].Value;
				if (!value.StartsWith("(") && !dictionary.ContainsKey(value))
				{
					dictionary.Add(value, value2);
				}
			}
			return dictionary;
		}
	}
}
