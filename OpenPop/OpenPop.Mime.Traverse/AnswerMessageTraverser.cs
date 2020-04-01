using System;
using System.Collections.Generic;

namespace OpenPop.Mime.Traverse
{
	public abstract class AnswerMessageTraverser<TAnswer> : IAnswerMessageTraverser<TAnswer>
	{
		public TAnswer VisitMessage(Message message)
		{
			if (message == null)
			{
				throw new ArgumentNullException("message");
			}
			return VisitMessagePart(message.MessagePart);
		}

		public TAnswer VisitMessagePart(MessagePart messagePart)
		{
			if (messagePart == null)
			{
				throw new ArgumentNullException("messagePart");
			}
			if (messagePart.IsMultiPart)
			{
				List<TAnswer> list = new List<TAnswer>(messagePart.MessageParts.Count);
				foreach (MessagePart messagePart2 in messagePart.MessageParts)
				{
					list.Add(VisitMessagePart(messagePart2));
				}
				return MergeLeafAnswers(list);
			}
			return CaseLeaf(messagePart);
		}

		protected abstract TAnswer CaseLeaf(MessagePart messagePart);

		protected abstract TAnswer MergeLeafAnswers(List<TAnswer> leafAnswers);
	}
}
