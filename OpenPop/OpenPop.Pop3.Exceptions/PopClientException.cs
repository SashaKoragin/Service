using System;
using System.Runtime.Serialization;

namespace OpenPop.Pop3.Exceptions
{
	[Serializable]
	public abstract class PopClientException : Exception
	{
		protected PopClientException()
		{
		}

		protected PopClientException(string message, Exception innerException)
			: base(message, innerException)
		{
			if (message == null)
			{
				throw new ArgumentNullException("message");
			}
			if (innerException == null)
			{
				throw new ArgumentNullException("innerException");
			}
		}

		protected PopClientException(string message)
			: base(message)
		{
			if (message == null)
			{
				throw new ArgumentNullException("message");
			}
		}

		protected PopClientException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}
