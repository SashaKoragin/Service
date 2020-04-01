using System;
using System.Collections.Generic;

namespace OpenPop.Mime.Traverse
{
	public abstract class MultipleMessagePartFinder : AnswerMessageTraverser<List<MessagePart>>
	{
		protected override List<MessagePart> MergeLeafAnswers(List<List<MessagePart>> leafAnswers)
		{
			if (leafAnswers == null)
			{
				throw new ArgumentNullException("leafAnswers");
			}
			List<MessagePart> list = new List<MessagePart>();
			foreach (List<MessagePart> leafAnswer in leafAnswers)
			{
				list.AddRange(leafAnswer);
			}
			return list;
		}
	}
}
