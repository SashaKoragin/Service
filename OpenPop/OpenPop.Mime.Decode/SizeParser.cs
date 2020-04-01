using System;
using System.Collections.Generic;
using System.Globalization;

namespace OpenPop.Mime.Decode
{
	internal static class SizeParser
	{
		private static readonly Dictionary<string, long> UnitsToMultiplicator = InitializeSizes();

		private static Dictionary<string, long> InitializeSizes()
		{
			Dictionary<string, long> dictionary = new Dictionary<string, long>();
			dictionary.Add("", 1L);
			dictionary.Add("B", 1L);
			dictionary.Add("KB", 1024L);
			dictionary.Add("MB", 1048576L);
			dictionary.Add("GB", 1073741824L);
			dictionary.Add("TB", 1099511627776L);
			return dictionary;
		}

		public static long Parse(string value)
		{
			value = value.Trim();
			string text = ExtractUnit(value);
			string s = value.Substring(0, value.Length - text.Length).Trim();
			long num = MultiplicatorForUnit(text);
			double num2 = double.Parse(s, NumberStyles.Number, CultureInfo.InvariantCulture);
			return (long)((double)num * num2);
		}

		private static string ExtractUnit(string sizeWithUnit)
		{
			int num = sizeWithUnit.Length - 1;
			int i;
			for (i = 0; i <= num && sizeWithUnit[num - i] != ' ' && !IsDigit(sizeWithUnit[num - i]); i++)
			{
			}
			return sizeWithUnit.Substring(sizeWithUnit.Length - i).ToUpperInvariant();
		}

		private static bool IsDigit(char value)
		{
			if (value >= '0')
			{
				return value <= '9';
			}
			return false;
		}

		private static long MultiplicatorForUnit(string unit)
		{
			unit = unit.ToUpperInvariant();
			if (!UnitsToMultiplicator.ContainsKey(unit))
			{
				throw new ArgumentException("illegal or unknown unit: \"" + unit + "\"", "unit");
			}
			return UnitsToMultiplicator[unit];
		}
	}
}
