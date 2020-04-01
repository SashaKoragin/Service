using OpenPop.Common;
using OpenPop.Mime.Decode;
using OpenPop.Mime.Header;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mime;
using System.Text;

namespace OpenPop.Mime
{
	public class MessagePart
	{
		private readonly IParsingErrorHandler _parsingErrorHandler;

		public ContentType ContentType
		{
			get;
			private set;
		}

		public string ContentDescription
		{
			get;
			private set;
		}

		public ContentTransferEncoding ContentTransferEncoding
		{
			get;
			private set;
		}

		public string ContentId
		{
			get;
			private set;
		}

		public ContentDisposition ContentDisposition
		{
			get;
			private set;
		}

		public Encoding BodyEncoding
		{
			get;
			private set;
		}

		public byte[] Body
		{
			get;
			private set;
		}

		public bool IsMultiPart => ContentType.MediaType.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase);

		public bool IsText
		{
			get
			{
				string mediaType = ContentType.MediaType;
				if (!mediaType.StartsWith("text/", StringComparison.OrdinalIgnoreCase))
				{
					return mediaType.Equals("message/rfc822", StringComparison.OrdinalIgnoreCase);
				}
				return true;
			}
		}

		public bool IsAttachment
		{
			get
			{
				if (IsText || IsMultiPart)
				{
					if (ContentDisposition != null)
					{
						return !ContentDisposition.Inline;
					}
					return false;
				}
				return true;
			}
		}

		public string FileName
		{
			get;
			private set;
		}

		public List<MessagePart> MessageParts
		{
			get;
			private set;
		}

		internal MessagePart(byte[] rawBody, MessageHeader headers, IParsingErrorHandler parsingErrorHandler = null)
		{
			if (rawBody == null)
			{
				throw new ArgumentNullException("rawBody");
			}
			if (headers == null)
			{
				throw new ArgumentNullException("headers");
			}
			_parsingErrorHandler = parsingErrorHandler;
			ContentType = headers.ContentType;
			ContentDescription = headers.ContentDescription;
			ContentTransferEncoding = headers.ContentTransferEncoding;
			ContentId = headers.ContentId;
			ContentDisposition = headers.ContentDisposition;
			FileName = FindFileName(ContentType, ContentDisposition, "(no name)");
			try
			{
				BodyEncoding = ParseBodyEncoding(ContentType.CharSet);
			}
			catch (FormatException exception)
			{
				BodyEncoding = Encoding.UTF8;
				if (_parsingErrorHandler == null)
				{
					throw;
				}
				_parsingErrorHandler.HandleParseError(new ParseError(exception, ContentType.CharSet, $"On body encoding. Encoding set to {BodyEncoding.EncodingName}"));
			}
			ParseBody(rawBody);
		}

		private static Encoding ParseBodyEncoding(string characterSet)
		{
			Encoding result = Encoding.ASCII;
			if (!string.IsNullOrEmpty(characterSet))
			{
				result = EncodingFinder.FindEncoding(characterSet);
			}
			return result;
		}

		private static string FindFileName(ContentType contentType, ContentDisposition contentDisposition, string defaultName)
		{
			if (contentType == null)
			{
				throw new ArgumentNullException("contentType");
			}
			if (contentDisposition != null && contentDisposition.FileName != null)
			{
				return contentDisposition.FileName;
			}
			if (contentType.Name != null)
			{
				return contentType.Name;
			}
			return defaultName;
		}

		private void ParseBody(byte[] rawBody)
		{
			if (IsMultiPart)
			{
				ParseMultiPartBody(rawBody);
			}
			else
			{
				Body = DecodeBody(rawBody, ContentTransferEncoding);
			}
		}

		private void ParseMultiPartBody(byte[] rawBody)
		{
			string boundary = ContentType.Boundary;
			List<byte[]> multiPartParts = GetMultiPartParts(rawBody, boundary);
			MessageParts = new List<MessagePart>(multiPartParts.Count);
			foreach (byte[] item in multiPartParts)
			{
				MessagePart messagePart = GetMessagePart(item);
				MessageParts.Add(messagePart);
			}
		}

		private static MessagePart GetMessagePart(byte[] rawMessageContent)
		{
			HeaderExtractor.ExtractHeadersAndBody(rawMessageContent, out MessageHeader headers, out byte[] body);
			return new MessagePart(body, headers);
		}

		private static List<byte[]> GetMultiPartParts(byte[] rawBody, string multipPartBoundary)
		{
			if (rawBody == null)
			{
				throw new ArgumentNullException("rawBody");
			}
			List<byte[]> list = new List<byte[]>();
			using (MemoryStream memoryStream = new MemoryStream(rawBody))
			{
				bool lastMultipartBoundaryFound;
				int num = FindPositionOfNextMultiPartBoundary(memoryStream, multipPartBoundary, out lastMultipartBoundaryFound) + ("--" + multipPartBoundary + "\r\n").Length;
				while (!lastMultipartBoundaryFound)
				{
					int num2 = FindPositionOfNextMultiPartBoundary(memoryStream, multipPartBoundary, out lastMultipartBoundaryFound) - "\r\n".Length;
					if (num2 <= -1)
					{
						num2 = (int)memoryStream.Length - "\r\n".Length;
						lastMultipartBoundaryFound = true;
						if (num >= num2)
						{
							return list;
						}
					}
					if (num >= num2)
					{
						num = num2 + ("\r\n--" + multipPartBoundary + "\r\n").Length;
					}
					else
					{
						int num3 = num2 - num;
						byte[] array = new byte[num3];
						Array.Copy(rawBody, num, array, 0, num3);
						list.Add(array);
						num = num2 + ("\r\n--" + multipPartBoundary + "\r\n").Length;
					}
				}
				return list;
			}
		}

		private static int FindPositionOfNextMultiPartBoundary(Stream stream, string multiPartBoundary, out bool lastMultipartBoundaryFound)
		{
			lastMultipartBoundaryFound = false;
			int result;
			string text;
			do
			{
				result = (int)stream.Position;
				text = StreamUtility.ReadLineAsAscii(stream);
				if (text == null)
				{
					return -1;
				}
			}
			while (!text.StartsWith("--" + multiPartBoundary, StringComparison.Ordinal));
			lastMultipartBoundaryFound = text.StartsWith("--" + multiPartBoundary + "--", StringComparison.OrdinalIgnoreCase);
			return result;
		}

		private static byte[] DecodeBody(byte[] messageBody, ContentTransferEncoding contentTransferEncoding)
		{
			if (messageBody == null)
			{
				throw new ArgumentNullException("messageBody");
			}
			switch (contentTransferEncoding)
			{
			case ContentTransferEncoding.QuotedPrintable:
				return QuotedPrintable.DecodeContentTransferEncoding(Encoding.ASCII.GetString(messageBody));
			case ContentTransferEncoding.Base64:
				return Base64.Decode(Encoding.ASCII.GetString(messageBody));
			case ContentTransferEncoding.SevenBit:
			case ContentTransferEncoding.EightBit:
			case ContentTransferEncoding.Binary:
				return messageBody;
			default:
				throw new ArgumentOutOfRangeException("contentTransferEncoding");
			}
		}

		public string GetBodyAsText()
		{
			return BodyEncoding.GetString(Body);
		}

		public void Save(FileInfo file)
		{
			if (file == null)
			{
				throw new ArgumentNullException("file");
			}
			using (FileStream messageStream = new FileStream(file.FullName, FileMode.Create))
			{
				Save(messageStream);
			}
		}

		public void Save(Stream messageStream)
		{
			if (messageStream == null)
			{
				throw new ArgumentNullException("messageStream");
			}
			messageStream.Write(Body, 0, Body.Length);
		}
	}
}
