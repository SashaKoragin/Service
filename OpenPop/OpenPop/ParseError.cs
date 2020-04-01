using System;

namespace OpenPop
{
	public class ParseError
	{
		public Exception Exception
		{
			get;
			private set;
		}

		public string FailedToBeParsed
		{
			get;
			private set;
		}

		public string Description
		{
			get;
			private set;
		}

		public ParseError(Exception exception, string failedToBeParsed, string description)
		{
			Exception = exception;
			FailedToBeParsed = failedToBeParsed;
			Description = description;
		}
	}
}
