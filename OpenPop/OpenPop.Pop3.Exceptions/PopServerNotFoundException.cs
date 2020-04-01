using System;
using System.Runtime.Serialization;

namespace OpenPop.Pop3.Exceptions
{
	[Serializable]
	public class PopServerNotFoundException : PopClientException
	{
		public PopServerNotFoundException(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		protected PopServerNotFoundException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}
