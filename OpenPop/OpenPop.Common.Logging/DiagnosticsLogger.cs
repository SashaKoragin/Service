#define TRACE
using System;
using System.Diagnostics;

namespace OpenPop.Common.Logging
{
	public class DiagnosticsLogger : ILog
	{
		public void LogError(string message)
		{
			if (message == null)
			{
				throw new ArgumentNullException("message");
			}
			Trace.WriteLine("OpenPOP: " + message);
		}

		public void LogDebug(string message)
		{
			if (message == null)
			{
				throw new ArgumentNullException("message");
			}
			Trace.WriteLine("OpenPOP: (DEBUG) " + message);
		}
	}
}
