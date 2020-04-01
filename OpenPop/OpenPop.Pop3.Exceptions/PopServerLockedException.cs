using System;
using System.Runtime.Serialization;

namespace OpenPop.Pop3.Exceptions
{
	[Serializable]
	public class PopServerLockedException : PopClientException
	{
		public PopServerLockedException(PopServerException innerException)
			: base("The account is locked or in use", innerException)
		{
		}

		protected PopServerLockedException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}
