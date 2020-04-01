using OpenPop.Common.Logging;
using OpenPop.Mime.Decode;
using System;
using System.Collections.Generic;
using System.Net.Mail;

namespace OpenPop.Mime.Header
{
	public class RfcMailAddress
	{
		public string Address
		{
			get;
			private set;
		}

		public string DisplayName
		{
			get;
			private set;
		}

		public string Raw
		{
			get;
			private set;
		}

		public MailAddress MailAddress
		{
			get;
			private set;
		}

		public bool HasValidMailAddress => MailAddress != null;

		private RfcMailAddress(MailAddress mailAddress, string raw)
		{
			if (mailAddress == null)
			{
				throw new ArgumentNullException("mailAddress");
			}
			if (raw == null)
			{
				throw new ArgumentNullException("raw");
			}
			MailAddress = mailAddress;
			Address = mailAddress.Address;
			DisplayName = mailAddress.DisplayName;
			Raw = raw;
		}

		private RfcMailAddress(string raw)
		{
			if (raw == null)
			{
				throw new ArgumentNullException("raw");
			}
			MailAddress = null;
			Address = string.Empty;
			DisplayName = raw;
			Raw = raw;
		}

		public override string ToString()
		{
			if (HasValidMailAddress)
			{
				return MailAddress.ToString();
			}
			return Raw;
		}

		internal static RfcMailAddress ParseMailAddress(string input)
		{
			if (input == null)
			{
				throw new ArgumentNullException("input");
			}
			input = EncodedWord.Decode(input.Trim());
			int num = input.LastIndexOf('<');
			int num2 = input.LastIndexOf('>');
			int num3 = num;
			int num4 = num2;
			while (num3 > 0 && input[num3 - 1] == '<' && input[num4 - 1] == '>')
			{
				num3--;
				num4--;
			}
			if (num3 != num)
			{
				input = input.Substring(0, num3) + input.Substring(num, num4 - num + 1);
			}
			int num5 = input.LastIndexOf('<');
			int num6 = input.LastIndexOf('>');
			try
			{
				if (num5 >= 0 && num6 >= 0)
				{
					string displayName = (num5 <= 0) ? string.Empty : input.Substring(0, num5).Trim();
					num5++;
					int length = num6 - num5;
					string text = input.Substring(num5, length).Trim();
					if (!string.IsNullOrEmpty(text))
					{
						return new RfcMailAddress(new MailAddress(text, displayName), input);
					}
				}
				if (input.Contains("@"))
				{
					return new RfcMailAddress(new MailAddress(input), input);
				}
			}
			catch (FormatException)
			{
				DefaultLogger.Log.LogError("RfcMailAddress: Improper mail address: \"" + input + "\"");
			}
			return new RfcMailAddress(input);
		}

		internal static List<RfcMailAddress> ParseMailAddresses(string input)
		{
			if (input == null)
			{
				throw new ArgumentNullException("input");
			}
			List<RfcMailAddress> list = new List<RfcMailAddress>();
			IEnumerable<string> enumerable = Utility.SplitStringWithCharNotInsideQuotes(input, ',');
			foreach (string item in enumerable)
			{
				list.Add(ParseMailAddress(item));
			}
			return list;
		}
	}
}
