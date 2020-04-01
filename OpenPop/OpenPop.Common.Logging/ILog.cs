namespace OpenPop.Common.Logging
{
	public interface ILog
	{
		void LogError(string message);

		void LogDebug(string message);
	}
}
