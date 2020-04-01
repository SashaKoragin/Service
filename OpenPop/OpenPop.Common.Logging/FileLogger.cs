using System;
using System.IO;

namespace OpenPop.Common.Logging
{
	public class FileLogger : ILog
	{
		private static readonly object LogLock;

		public static bool Enabled
		{
			get;
			set;
		}

		public static bool Verbose
		{
			get;
			set;
		}

		public static FileInfo LogFile
		{
			get;
			set;
		}

		static FileLogger()
		{
			LogFile = new FileInfo("OpenPOP.log");
			Enabled = true;
			Verbose = false;
			LogLock = new object();
		}

		private static void LogToFile(string text)
		{
			if (text == null)
			{
				throw new ArgumentNullException("text");
			}
			lock (LogLock)
			{
				using (StreamWriter streamWriter = LogFile.AppendText())
				{
					streamWriter.WriteLine(DateTime.Now + " " + text);
					streamWriter.Flush();
				}
			}
		}

		public void LogError(string message)
		{
			if (Enabled)
			{
				LogToFile(message);
			}
		}

		public void LogDebug(string message)
		{
			if (Enabled && Verbose)
			{
				LogToFile("DEBUG: " + message);
			}
		}
	}
}
