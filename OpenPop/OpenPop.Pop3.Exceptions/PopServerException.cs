using System;
using System.Runtime.Serialization;

namespace OpenPop.Pop3.Exceptions
{
	[Serializable]
	public class PopServerException : PopClientException
	{
		public PopServerException(string message)
			: base(message)
		{
		}

		protected PopServerException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}
