using System;

namespace OpenPop.Common.Logging
{
	public static class DefaultLogger
	{
		public static ILog Log
		{
			get;
			private set;
		}

		static DefaultLogger()
		{
			Log = new DiagnosticsLogger();
		}

		public static void SetLog(ILog newLogger)
		{
			if (newLogger == null)
			{
				throw new ArgumentNullException("newLogger");
			}
			Log = newLogger;
		}
	}
}
