using System;
using System.Runtime.Serialization;

namespace OpenPop.Pop3.Exceptions
{
	[Serializable]
	public class InvalidLoginException : PopClientException
	{
		public InvalidLoginException(Exception innerException)
			: base("Server did not accept user credentials", innerException)
		{
		}

		protected InvalidLoginException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}
