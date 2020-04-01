using System;

namespace OpenPop.Mime.Traverse
{
	internal class FindFirstMessagePartWithMediaType : IQuestionAnswerMessageTraverser<string, MessagePart>
	{
		public MessagePart VisitMessage(Message message, string question)
		{
			if (message == null)
			{
				throw new ArgumentNullException("message");
			}
			return VisitMessagePart(message.MessagePart, question);
		}

		public MessagePart VisitMessagePart(MessagePart messagePart, string question)
		{
			if (messagePart == null)
			{
				throw new ArgumentNullException("messagePart");
			}
			if (messagePart.ContentType.MediaType.Equals(question, StringComparison.OrdinalIgnoreCase))
			{
				return messagePart;
			}
			if (messagePart.IsMultiPart)
			{
				foreach (MessagePart messagePart3 in messagePart.MessageParts)
				{
					MessagePart messagePart2 = VisitMessagePart(messagePart3, question);
					if (messagePart2 != null)
					{
						return messagePart2;
					}
				}
			}
			return null;
		}
	}
}
