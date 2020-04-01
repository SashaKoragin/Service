using System;
using System.Runtime.Serialization;

namespace OpenPop.Pop3.Exceptions
{
	[Serializable]
	public class PopServerNotAvailableException : PopClientException
	{
		public PopServerNotAvailableException(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		protected PopServerNotAvailableException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}
