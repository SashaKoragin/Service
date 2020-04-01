using OpenPop.Mime.Decode;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Net.Mail;
using System.Net.Mime;

namespace OpenPop.Mime.Header
{
	public sealed class MessageHeader
	{
		private readonly IParsingErrorHandler _parsingErrorHandler;

		public NameValueCollection UnknownHeaders
		{
			get;
			private set;
		}

		public string ContentDescription
		{
			get;
			private set;
		}

		public string ContentId
		{
			get;
			private set;
		}

		public List<string> Keywords
		{
			get;
			private set;
		}

		public List<RfcMailAddress> DispositionNotificationTo
		{
			get;
			private set;
		}

		public List<Received> Received
		{
			get;
			private set;
		}

		public MailPriority Importance
		{
			get;
			private set;
		}

		public ContentTransferEncoding ContentTransferEncoding
		{
			get;
			private set;
		}

		public List<RfcMailAddress> Cc
		{
			get;
			private set;
		}

		public List<RfcMailAddress> Bcc
		{
			get;
			private set;
		}

		public List<RfcMailAddress> To
		{
			get;
			private set;
		}

		public RfcMailAddress From
		{
			get;
			private set;
		}

		public RfcMailAddress ReplyTo
		{
			get;
			private set;
		}

		public List<string> InReplyTo
		{
			get;
			private set;
		}

		public List<string> References
		{
			get;
			private set;
		}

		public RfcMailAddress Sender
		{
			get;
			private set;
		}

		public ContentType ContentType
		{
			get;
			private set;
		}

		public ContentDisposition ContentDisposition
		{
			get;
			private set;
		}

		public string Date
		{
			get;
			private set;
		}

		public DateTime DateSent
		{
			get;
			private set;
		}

		public string MessageId
		{
			get;
			private set;
		}

		public string MimeVersion
		{
			get;
			private set;
		}

		public RfcMailAddress ReturnPath
		{
			get;
			private set;
		}

		public string Subject
		{
			get;
			private set;
		}

		internal MessageHeader(NameValueCollection headers, IParsingErrorHandler parsingErrorHandler = null)
		{
			_parsingErrorHandler = parsingErrorHandler;
			if (headers == null)
			{
				throw new ArgumentNullException("headers");
			}
			To = new List<RfcMailAddress>(0);
			Cc = new List<RfcMailAddress>(0);
			Bcc = new List<RfcMailAddress>(0);
			Received = new List<Received>();
			Keywords = new List<string>();
			InReplyTo = new List<string>(0);
			References = new List<string>(0);
			DispositionNotificationTo = new List<RfcMailAddress>();
			UnknownHeaders = new NameValueCollection();
			Importance = MailPriority.Normal;
			ContentTransferEncoding = ContentTransferEncoding.SevenBit;
			ContentType = new ContentType("text/plain; charset=us-ascii");
			ParseHeaders(headers);
		}

		private void ParseHeaders(NameValueCollection headers)
		{
			if (headers == null)
			{
				throw new ArgumentNullException("headers");
			}
			foreach (string key in headers.Keys)
			{
				string[] values = headers.GetValues(key);
				if (values != null)
				{
					string[] array = values;
					foreach (string text2 in array)
					{
						try
						{
							ParseHeader(key, text2);
						}
						catch (Exception exception)
						{
							if (_parsingErrorHandler == null)
							{
								throw;
							}
							_parsingErrorHandler.HandleParseError(new ParseError(exception, text2, $"On header {key}"));
						}
					}
				}
			}
		}

		private void ParseHeader(string headerName, string headerValue)
		{
			if (headerName == null)
			{
				throw new ArgumentNullException("headerName");
			}
			if (headerValue == null)
			{
				throw new ArgumentNullException("headerValue");
			}
			switch (headerName.ToUpperInvariant())
			{
			case "TO":
				To = RfcMailAddress.ParseMailAddresses(headerValue);
				break;
			case "CC":
				Cc = RfcMailAddress.ParseMailAddresses(headerValue);
				break;
			case "BCC":
				Bcc = RfcMailAddress.ParseMailAddresses(headerValue);
				break;
			case "FROM":
				From = RfcMailAddress.ParseMailAddress(headerValue);
				break;
			case "REPLY-TO":
				ReplyTo = RfcMailAddress.ParseMailAddress(headerValue);
				break;
			case "SENDER":
				Sender = RfcMailAddress.ParseMailAddress(headerValue);
				break;
			case "KEYWORDS":
			{
				string[] array = headerValue.Split(',');
				string[] array2 = array;
				foreach (string text in array2)
				{
					Keywords.Add(Utility.RemoveQuotesIfAny(text.Trim()));
				}
				break;
			}
			case "RECEIVED":
				Received.Add(new Received(headerValue.Trim()));
				break;
			case "IMPORTANCE":
				Importance = HeaderFieldParser.ParseImportance(headerValue.Trim());
				break;
			case "DISPOSITION-NOTIFICATION-TO":
				DispositionNotificationTo = RfcMailAddress.ParseMailAddresses(headerValue);
				break;
			case "MIME-VERSION":
				MimeVersion = headerValue.Trim();
				break;
			case "SUBJECT":
				Subject = EncodedWord.Decode(headerValue);
				break;
			case "RETURN-PATH":
				ReturnPath = RfcMailAddress.ParseMailAddress(headerValue);
				break;
			case "MESSAGE-ID":
				MessageId = HeaderFieldParser.ParseId(headerValue);
				break;
			case "IN-REPLY-TO":
				InReplyTo = HeaderFieldParser.ParseMultipleIDs(headerValue);
				break;
			case "REFERENCES":
				References = HeaderFieldParser.ParseMultipleIDs(headerValue);
				break;
			case "DATE":
				Date = headerValue.Trim();
				DateSent = Rfc2822DateTime.StringToDate(headerValue);
				break;
			case "CONTENT-TRANSFER-ENCODING":
				ContentTransferEncoding = HeaderFieldParser.ParseContentTransferEncoding(headerValue.Trim());
				break;
			case "CONTENT-DESCRIPTION":
				ContentDescription = EncodedWord.Decode(headerValue.Trim());
				break;
			case "CONTENT-TYPE":
				ContentType = HeaderFieldParser.ParseContentType(headerValue);
				break;
			case "CONTENT-DISPOSITION":
				ContentDisposition = HeaderFieldParser.ParseContentDisposition(headerValue);
				break;
			case "CONTENT-ID":
				ContentId = HeaderFieldParser.ParseId(headerValue);
				break;
			default:
				UnknownHeaders.Add(headerName, headerValue);
				break;
			}
		}
	}
}
