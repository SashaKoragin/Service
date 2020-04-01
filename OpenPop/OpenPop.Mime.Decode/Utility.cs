using System;
using System.Collections.Generic;

namespace OpenPop.Mime.Decode
{
	internal static class Utility
	{
		public static string RemoveQuotesIfAny(string text)
		{
			if (text == null)
			{
				throw new ArgumentNullException("text");
			}
			if (text.Length > 1 && text[0] == '"' && text[text.Length - 1] == '"')
			{
				return text.Substring(1, text.Length - 2);
			}
			return text;
		}

		public static List<string> SplitStringWithCharNotInsideQuotes(string input, char toSplitAt)
		{
			List<string> list = new List<string>();
			int num = 0;
			bool flag = false;
			char[] array = input.ToCharArray();
			for (int i = 0; i < array.Length; i++)
			{
				char c = array[i];
				if (c == '"')
				{
					flag = !flag;
				}
				if (c == toSplitAt && !flag)
				{
					int length = i - num;
					list.Add(input.Substring(num, length));
					num = i + 1;
				}
			}
			list.Add(input.Substring(num, input.Length - num));
			return list;
		}
	}
}
