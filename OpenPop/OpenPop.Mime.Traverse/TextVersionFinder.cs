using System;
using System.Collections.Generic;

namespace OpenPop.Mime.Traverse
{
	internal class TextVersionFinder : MultipleMessagePartFinder
	{
		protected override List<MessagePart> CaseLeaf(MessagePart messagePart)
		{
			if (messagePart == null)
			{
				throw new ArgumentNullException("messagePart");
			}
			List<MessagePart> list = new List<MessagePart>(1);
			if (messagePart.IsText)
			{
				list.Add(messagePart);
			}
			return list;
		}
	}
}
