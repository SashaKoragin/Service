using System;
using System.Runtime.Serialization;

namespace OpenPop.Pop3.Exceptions
{
	[Serializable]
	public class LoginDelayException : PopClientException
	{
		public LoginDelayException(PopServerException innerException)
			: base("Login denied because of recent connection to this maildrop. Increase time between connections.", innerException)
		{
		}

		protected LoginDelayException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}
