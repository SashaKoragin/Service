using System;
using System.Collections.Generic;

namespace OpenPop.Mime.Traverse
{
	internal class FindAllMessagePartsWithMediaType : IQuestionAnswerMessageTraverser<string, List<MessagePart>>
	{
		public List<MessagePart> VisitMessage(Message message, string question)
		{
			if (message == null)
			{
				throw new ArgumentNullException("message");
			}
			return VisitMessagePart(message.MessagePart, question);
		}

		public List<MessagePart> VisitMessagePart(MessagePart messagePart, string question)
		{
			if (messagePart == null)
			{
				throw new ArgumentNullException("messagePart");
			}
			List<MessagePart> list = new List<MessagePart>();
			if (messagePart.ContentType.MediaType.Equals(question, StringComparison.OrdinalIgnoreCase))
			{
				list.Add(messagePart);
			}
			if (messagePart.IsMultiPart)
			{
				foreach (MessagePart messagePart2 in messagePart.MessageParts)
				{
					List<MessagePart> collection = VisitMessagePart(messagePart2, question);
					list.AddRange(collection);
				}
				return list;
			}
			return list;
		}
	}
}
