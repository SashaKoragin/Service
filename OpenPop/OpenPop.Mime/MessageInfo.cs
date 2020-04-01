namespace OpenPop.Mime
{
	public class MessageInfo
	{
		private readonly int _number;

		private readonly string _identifier;

		private readonly int _size;

		public int Number => _number;

		public string Identifier => _identifier;

		public int Size => _size;

		public MessageInfo(int number, string id, int size)
		{
			_number = number;
			_identifier = id;
			_size = size;
		}

		public override string ToString()
		{
			return string.Format("Message #{0} ({2} octets): '{1}'", Number, Identifier, Size);
		}
	}
}
