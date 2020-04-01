using System;
using System.Runtime.Serialization;

namespace OpenPop.Pop3.Exceptions
{
	[Serializable]
	public class InvalidUseException : PopClientException
	{
		public InvalidUseException(string message)
			: base(message)
		{
		}

		protected InvalidUseException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}
