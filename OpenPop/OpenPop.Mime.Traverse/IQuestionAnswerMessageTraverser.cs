namespace OpenPop.Mime.Traverse
{
	public interface IQuestionAnswerMessageTraverser<TQuestion, TAnswer>
	{
		TAnswer VisitMessage(Message message, TQuestion question);

		TAnswer VisitMessagePart(MessagePart messagePart, TQuestion question);
	}
}
