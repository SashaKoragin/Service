using OpenPop.Common.Logging;
using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OpenPop.Mime.Decode
{
	public static class Rfc2822DateTime
	{
		private const string REGEX_OLD_TIMEZONE_FORMATS = "UT|GMT|EST|EDT|CST|CDT|MST|MDT|PST|MSK|PDT|[A-I]|[K-Y]|Z";

		private const string REGEX_NEW_TIMEZONE_FORMATS = "[\\+-](?<hours>\\d\\d)(?<minutes>\\d\\d)";

		public static string[] CustomDateTimeFormats
		{
			private get;
			set;
		}

		public static DateTime StringToDate(string inputDate)
		{
			if (inputDate == null)
			{
				throw new ArgumentNullException("inputDate");
			}
			inputDate = FixSpecialCases(inputDate);
			inputDate = StripCommentsAndExcessWhitespace(inputDate);
			try
			{
				DateTime dateTime = ExtractDateTime(inputDate);
				if (dateTime == DateTime.MinValue)
				{
					return dateTime;
				}
				ValidateDayNameIfAny(dateTime, inputDate);
				dateTime = new DateTime(dateTime.Ticks, DateTimeKind.Utc);
				return AdjustTimezone(dateTime, inputDate);
			}
			catch (FormatException ex)
			{
				throw new ArgumentException("Could not parse date: " + ex.Message + ". Input was: \"" + inputDate + "\"", ex);
			}
			catch (ArgumentException ex2)
			{
				throw new ArgumentException("Could not parse date: " + ex2.Message + ". Input was: \"" + inputDate + "\"", ex2);
			}
		}

		private static DateTime AdjustTimezone(DateTime dateTime, string dateInput)
		{
			string[] array = dateInput.Split(' ');
			string input = array[array.Length - 1];
			input = Regex.Replace(input, "UT|GMT|EST|EDT|CST|CDT|MST|MDT|PST|MSK|PDT|[A-I]|[K-Y]|Z", MatchEvaluator);
			Match match = Regex.Match(input, "[\\+-](?<hours>\\d\\d)(?<minutes>\\d\\d)");
			if (match.Success)
			{
				int num = int.Parse(match.Groups["hours"].Value);
				int num2 = int.Parse(match.Groups["minutes"].Value);
				int num3 = (match.Value[0] != '+') ? 1 : (-1);
				dateTime = dateTime.AddHours(num3 * num);
				dateTime = dateTime.AddMinutes(num3 * num2);
				return dateTime;
			}
			DefaultLogger.Log.LogDebug("No timezone found in date: " + dateInput + ". Using -0000 as default.");
			return dateTime;
		}

		private static string MatchEvaluator(Match match)
		{
			if (!match.Success)
			{
				throw new ArgumentException("Match success are always true");
			}
			switch (match.Value)
			{
			case "A":
				return "+0100";
			case "B":
				return "+0200";
			case "C":
				return "+0300";
			case "D":
				return "+0400";
			case "E":
				return "+0500";
			case "F":
				return "+0600";
			case "G":
				return "+0700";
			case "H":
				return "+0800";
			case "I":
				return "+0900";
			case "K":
				return "+1000";
			case "L":
				return "+1100";
			case "M":
				return "+1200";
			case "N":
				return "-0100";
			case "O":
				return "-0200";
			case "P":
				return "-0300";
			case "Q":
				return "-0400";
			case "R":
				return "-0500";
			case "S":
				return "-0600";
			case "T":
				return "-0700";
			case "U":
				return "-0800";
			case "V":
				return "-0900";
			case "W":
				return "-1000";
			case "X":
				return "-1100";
			case "Y":
				return "-1200";
			case "Z":
			case "UT":
			case "GMT":
				return "+0000";
			case "EDT":
				return "-0400";
			case "EST":
				return "-0500";
			case "CDT":
				return "-0500";
			case "CST":
				return "-0600";
			case "MDT":
				return "-0600";
			case "MST":
				return "-0700";
			case "PDT":
				return "-0700";
			case "PST":
				return "-0800";
			case "MSK":
				return "+0400";
			default:
				throw new ArgumentException("Unexpected input");
			}
		}

		private static DateTime ExtractDateTime(string dateInput)
		{
			if (dateInput == null)
			{
				throw new ArgumentNullException("dateInput");
			}
			if (CustomDateTimeFormats != null && CustomDateTimeFormats.Length > 0)
			{
				string text = dateInput.Trim();
				if (text.Contains(" "))
				{
					string[] array = text.Split(' ');
					string input = array[array.Length - 1];
					input = Regex.Replace(input, "UT|GMT|EST|EDT|CST|CDT|MST|MDT|PST|MSK|PDT|[A-I]|[K-Y]|Z", MatchEvaluator);
					Match match = Regex.Match(input, "[\\+-](?<hours>\\d\\d)(?<minutes>\\d\\d)");
					if (match.Success)
					{
						text = text.Substring(0, text.Length - array[array.Length - 1].Length).Trim();
					}
				}
				try
				{
					DateTime result = DateTime.ParseExact(text, CustomDateTimeFormats, null, DateTimeStyles.None);
					DefaultLogger.Log.LogDebug($"Successfully parsed date input \"{dateInput}\" using a custom format. Converted to date: {result.ToString()}");
					return result;
				}
				catch (FormatException)
				{
				}
			}
			Match match2 = Regex.Match(dateInput, "(\\d\\d? .+ (\\d\\d\\d\\d|\\d\\d) \\d?\\d:\\d?\\d(:\\d?\\d)?)|((\\d\\d\\d\\d|\\d\\d)-\\d?\\d-\\d?\\d \\d?\\d:\\d?\\d(:\\d?\\d)?)|(\\d\\d?-[A-Za-z]{3}-(\\d\\d\\d\\d|\\d\\d) \\d?\\d:\\d?\\d(:\\d?\\d)?)");
			if (match2.Success)
			{
				try
				{
					return Convert.ToDateTime(match2.Value, CultureInfo.InvariantCulture);
				}
				catch (FormatException)
				{
					DefaultLogger.Log.LogError("The given date appeared to be in a valid format, but could not be converted to a DateTime object: " + dateInput);
				}
			}
			else
			{
				DefaultLogger.Log.LogError("The given date does not appear to be in a valid format: " + dateInput);
			}
			return DateTime.MinValue;
		}

		private static void ValidateDayNameIfAny(DateTime dateTime, string dateInput)
		{
			if (dateInput.Length >= 4 && dateInput[3] == ',')
			{
				string text = dateInput.Substring(0, 3);
				if ((dateTime.DayOfWeek == DayOfWeek.Monday && !text.Equals("Mon")) || (dateTime.DayOfWeek == DayOfWeek.Tuesday && !text.Equals("Tue")) || (dateTime.DayOfWeek == DayOfWeek.Wednesday && !text.Equals("Wed")) || (dateTime.DayOfWeek == DayOfWeek.Thursday && !text.Equals("Thu")) || (dateTime.DayOfWeek == DayOfWeek.Friday && !text.Equals("Fri")) || (dateTime.DayOfWeek == DayOfWeek.Saturday && !text.Equals("Sat")) || (dateTime.DayOfWeek == DayOfWeek.Sunday && !text.Equals("Sun")))
				{
					DefaultLogger.Log.LogDebug("Day-name does not correspond to the weekday of the date: " + dateInput);
				}
			}
		}

		private static string StripCommentsAndExcessWhitespace(string input)
		{
			if (input == null)
			{
				throw new ArgumentNullException("input");
			}
			input = Regex.Replace(input, "(\\((?>\\((?<C>)|\\)(?<-C>)|.?)*(?(C)(?!))\\))", "");
			input = Regex.Replace(input, "\\s+", " ");
			input = Regex.Replace(input, "^\\s+", "");
			input = Regex.Replace(input, "\\s+$", "");
			input = Regex.Replace(input, " ?: ?", ":");
			return input;
		}

		private static string FixSpecialCases(string inputDate)
		{
			string pattern = string.Format("{0} +{1} +{2} +{3} +{4} +\\({5}\\)", "(?<weekDay>Mon|Tue|Wed|Thu|Fri|Sat|Sun)", "(?<month>[A-Za-z]+)", "(?<day>\\d?\\d)", "(?<time>\\d?\\d:\\d?\\d(:\\d?\\d)?)", "(?<year>\\d\\d\\d\\d)", "(?<timeZone>[A-Z]{3})");
			Match match = Regex.Match(inputDate, pattern);
			if (match.Success)
			{
				Group group = match.Groups["month"];
				Group group2 = match.Groups["day"];
				Group group3 = match.Groups["year"];
				Group group4 = match.Groups["time"];
				Group group5 = match.Groups["timeZone"];
				return $"{group2} {group} {group3} {group4} {group5}";
			}
			return inputDate;
		}
	}
}
