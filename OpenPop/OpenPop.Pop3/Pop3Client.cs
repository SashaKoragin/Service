using OpenPop.Common;
using OpenPop.Common.Logging;
using OpenPop.Mime;
using OpenPop.Mime.Header;
using OpenPop.Pop3.Exceptions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Security;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;

namespace OpenPop.Pop3
{
	public class Pop3Client : Disposable
	{
		private Stream Stream
		{
			get;
			set;
		}

		private string LastServerResponse
		{
			get;
			set;
		}

		private string ApopTimeStamp
		{
			get;
			set;
		}

		private ConnectionState State
		{
			get;
			set;
		}

		public bool Connected
		{
			get;
			private set;
		}

		public bool ApopSupported
		{
			get;
			private set;
		}

		public Pop3Client()
		{
			SetInitialValues();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && !base.IsDisposed && Connected)
			{
				try
				{
					Disconnect();
				}
				catch
				{
				}
			}
			base.Dispose(disposing);
		}

		public void Connect(Stream stream)
		{
			AssertDisposed();
			if (State != 0)
			{
				throw new InvalidUseException("You cannot ask to connect to a POP3 server, when we are already connected to one. Disconnect first.");
			}
			if (stream == null)
			{
				throw new ArgumentNullException("stream");
			}
			Stream = stream;
			string text = StreamUtility.ReadLineAsAscii(Stream);
			try
			{
				State = ConnectionState.Authorization;
				DefaultLogger.Log.LogDebug($"Connect-Response: \"{text}\"");
				IsOkResponse(text);
				ExtractApopTimestamp(text);
				Connected = true;
			}
			catch (PopServerException innerException)
			{
				DisconnectStreams();
				DefaultLogger.Log.LogError("Connect(): Error with connection, maybe POP3 server not exist");
				DefaultLogger.Log.LogDebug("Last response from server was: " + LastServerResponse);
				throw new PopServerNotAvailableException("Server is not available", innerException);
			}
		}

		public void Connect(string hostname, int port, bool useSsl)
		{
			Connect(hostname, port, useSsl, 60000, 60000, null);
		}

		public void Connect(string hostname, int port, bool useSsl, int receiveTimeout, int sendTimeout, RemoteCertificateValidationCallback certificateValidator)
		{
			AssertDisposed();
			if (hostname == null)
			{
				throw new ArgumentNullException("hostname");
			}
			if (hostname.Length == 0)
			{
				throw new ArgumentException("hostname cannot be empty", "hostname");
			}
			if (port > 65535 || port < 0)
			{
				throw new ArgumentOutOfRangeException("port");
			}
			if (receiveTimeout < 0)
			{
				throw new ArgumentOutOfRangeException("receiveTimeout");
			}
			if (sendTimeout < 0)
			{
				throw new ArgumentOutOfRangeException("sendTimeout");
			}
			if (State != 0)
			{
				throw new InvalidUseException("You cannot ask to connect to a POP3 server, when we are already connected to one. Disconnect first.");
			}
			TcpClient tcpClient = new TcpClient();
			tcpClient.ReceiveTimeout = receiveTimeout;
			tcpClient.SendTimeout = sendTimeout;
			try
			{
				tcpClient.Connect(hostname, port);
			}
			catch (SocketException ex)
			{
				tcpClient.Close();
				DefaultLogger.Log.LogError("Connect(): " + ex.Message);
				throw new PopServerNotFoundException("Server not found", ex);
			}
			Stream stream;
			if (useSsl)
			{
				SslStream sslStream = (certificateValidator != null) ? new SslStream(tcpClient.GetStream(), leaveInnerStreamOpen: false, certificateValidator) : new SslStream(tcpClient.GetStream(), leaveInnerStreamOpen: false);
				sslStream.ReadTimeout = receiveTimeout;
				sslStream.WriteTimeout = sendTimeout;
				System.Security.Cryptography.X509Certificates.X509Certificate2Collection xc = new System.Security.Cryptography.X509Certificates.X509Certificate2Collection();
                sslStream.AuthenticateAsClient(hostname, xc, System.Security.Authentication.SslProtocols.Tls12, false);
                stream = sslStream;
			}
			else
			{
				stream = tcpClient.GetStream();
			}
			Connect(stream);
		}

		public void Disconnect()
		{
			AssertDisposed();
			if (State == ConnectionState.Disconnected)
			{
				throw new InvalidUseException("You cannot disconnect a connection which is already disconnected");
			}
			try
			{
				SendCommand("QUIT");
			}
			finally
			{
				DisconnectStreams();
			}
		}

		public void Authenticate(string username, string password)
		{
			AssertDisposed();
			Authenticate(username, password, AuthenticationMethod.Auto);
		}

		public void Authenticate(string username, string password, AuthenticationMethod authenticationMethod)
		{
			AssertDisposed();
			if (username == null)
			{
				throw new ArgumentNullException("username");
			}
			if (password == null)
			{
				throw new ArgumentNullException("password");
			}
			if (State != ConnectionState.Authorization)
			{
				throw new InvalidUseException("You have to be connected and not authorized when trying to authorize yourself");
			}
			try
			{
				switch (authenticationMethod)
				{
				case AuthenticationMethod.UsernameAndPassword:
					AuthenticateUsingUserAndPassword(username, password);
					break;
				case AuthenticationMethod.Apop:
					AuthenticateUsingApop(username, password);
					break;
				case AuthenticationMethod.Auto:
					if (ApopSupported)
					{
						AuthenticateUsingApop(username, password);
					}
					else
					{
						AuthenticateUsingUserAndPassword(username, password);
					}
					break;
				case AuthenticationMethod.CramMd5:
					AuthenticateUsingCramMd5(username, password);
					break;
				}
			}
			catch (PopServerException ex)
			{
				DefaultLogger.Log.LogError("Problem logging in using method " + authenticationMethod + ". Server response was: " + LastServerResponse);
				CheckFailedLoginServerResponse(LastServerResponse, ex);
				throw new InvalidLoginException(ex);
			}
			State = ConnectionState.Transaction;
		}

		private void AuthenticateUsingUserAndPassword(string username, string password)
		{
			SendCommand("USER " + username);
			SendCommand("PASS " + password);
		}

		private void AuthenticateUsingApop(string username, string password)
		{
			if (!ApopSupported)
			{
				throw new NotSupportedException("APOP is not supported on this server");
			}
			SendCommand("APOP " + username + " " + Apop.ComputeDigest(password, ApopTimeStamp));
		}

		private void AuthenticateUsingCramMd5(string username, string password)
		{
			try
			{
				SendCommand("AUTH CRAM-MD5");
			}
			catch (PopServerException innerException)
			{
				throw new NotSupportedException("CRAM-MD5 authentication not supported", innerException);
			}
			string challenge = LastServerResponse.Substring(2);
			string command = CramMd5.ComputeDigest(username, password, challenge);
			SendCommand(command);
		}

		public int GetMessageCount()
		{
			AssertDisposed();
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("You cannot get the message count without authenticating yourself towards the server first");
			}
			return SendCommandIntResponse("STAT", 1);
		}

		public void DeleteMessage(int messageNumber)
		{
			AssertDisposed();
			ValidateMessageNumber(messageNumber);
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("You cannot delete any messages without authenticating yourself towards the server first");
			}
			SendCommand("DELE " + messageNumber);
		}

		public void DeleteAllMessages()
		{
			AssertDisposed();
			int messageCount = GetMessageCount();
			for (int num = messageCount; num > 0; num--)
			{
				DeleteMessage(num);
			}
		}

		public void NoOperation()
		{
			AssertDisposed();
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("You cannot use the NOOP command unless you are authenticated to the server");
			}
			SendCommand("NOOP");
		}

		public void Reset()
		{
			AssertDisposed();
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("You cannot use the RSET command unless you are authenticated to the server");
			}
			SendCommand("RSET");
		}

		public string GetMessageUid(int messageNumber)
		{
			AssertDisposed();
			ValidateMessageNumber(messageNumber);
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("Cannot get message ID, when the user has not been authenticated yet");
			}
			SendCommand("UIDL " + messageNumber);
			return LastServerResponse.Split(' ')[2];
		}

		public List<string> GetMessageUids()
		{
			AssertDisposed();
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("Cannot get message IDs, when the user has not been authenticated yet");
			}
			SendCommand("UIDL");
			List<string> list = new List<string>();
			string text;
			while (!IsLastLineInMultiLineResponse(text = StreamUtility.ReadLineAsAscii(Stream)))
			{
				list.Add(text.Split(' ')[1]);
			}
			return list;
		}

		public int GetMessageSize(int messageNumber)
		{
			AssertDisposed();
			ValidateMessageNumber(messageNumber);
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("Cannot get message size, when the user has not been authenticated yet");
			}
			return SendCommandIntResponse("LIST " + messageNumber, 2);
		}

		public List<int> GetMessageSizes()
		{
			AssertDisposed();
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("Cannot get message sizes, when the user has not been authenticated yet");
			}
			SendCommand("LIST");
			List<int> list = new List<int>();
			string text;
			while (!".".Equals(text = StreamUtility.ReadLineAsAscii(Stream)))
			{
				list.Add(int.Parse(text.Split(' ')[1], CultureInfo.InvariantCulture));
			}
			return list;
		}

		public List<MessageInfo> GetMessageInfos()
		{
			AssertDisposed();
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("Cannot get message infos, when the user has not been authenticated yet");
			}
			SendCommand("UIDL");
			Dictionary<int, string> dictionary = new Dictionary<int, string>();
			string text;
			while (!IsLastLineInMultiLineResponse(text = StreamUtility.ReadLineAsAscii(Stream)))
			{
				string[] array = text.Split(' ');
				int key = int.Parse(array[0], CultureInfo.InvariantCulture);
				string value = array[1];
				dictionary.Add(key, value);
			}
			SendCommand("LIST");
			Dictionary<int, int> dictionary2 = new Dictionary<int, int>();
			string text2;
			while (!IsLastLineInMultiLineResponse(text2 = StreamUtility.ReadLineAsAscii(Stream)))
			{
				string[] array2 = text2.Split(' ');
				int key2 = int.Parse(array2[0], CultureInfo.InvariantCulture);
				int value2 = int.Parse(array2[1], CultureInfo.InvariantCulture);
				dictionary2.Add(key2, value2);
			}
			if (dictionary2.Count != dictionary.Count)
			{
				throw new PopServerException("Server LIST and UIDL responses do not match.");
			}
			int count = dictionary.Count;
			List<MessageInfo> list = new List<MessageInfo>(count);
			foreach (int key3 in dictionary.Keys)
			{
				string id = dictionary[key3];
				int size = dictionary2[key3];
				list.Add(new MessageInfo(key3, id, size));
			}
			return list;
		}

		public Message GetMessage(int messageNumber, IParsingErrorHandler parsingErrorHandler = null)
		{
			AssertDisposed();
			ValidateMessageNumber(messageNumber);
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("Cannot fetch a message, when the user has not been authenticated yet");
			}
			byte[] messageAsBytes = GetMessageAsBytes(messageNumber);
			return new Message(messageAsBytes, parsingErrorHandler);
		}

		public byte[] GetMessageAsBytes(int messageNumber)
		{
			AssertDisposed();
			ValidateMessageNumber(messageNumber);
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("Cannot fetch a message, when the user has not been authenticated yet");
			}
			return GetMessageAsBytes(messageNumber, askOnlyForHeaders: false);
		}

		public MessageHeader GetMessageHeaders(int messageNumber)
		{
			AssertDisposed();
			ValidateMessageNumber(messageNumber);
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("Cannot fetch a message, when the user has not been authenticated yet");
			}
			byte[] messageAsBytes = GetMessageAsBytes(messageNumber, askOnlyForHeaders: true);
			return new Message(messageAsBytes, parseBody: false).Headers;
		}

		public Dictionary<string, List<string>> Capabilities()
		{
			AssertDisposed();
			if (State != ConnectionState.Authorization && State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("Capability command only available while connected or authenticated");
			}
			SendCommand("CAPA");
			Dictionary<string, List<string>> dictionary = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
			string text;
			while (!IsLastLineInMultiLineResponse(text = StreamUtility.ReadLineAsAscii(Stream)))
			{
				string[] array = text.Split(' ');
				string key = array[0];
				List<string> list = new List<string>();
				for (int i = 1; i < array.Length; i++)
				{
					list.Add(array[i]);
				}
				dictionary.Add(key, list);
			}
			return dictionary;
		}

		private void ExtractApopTimestamp(string response)
		{
			if (response == null)
			{
				throw new ArgumentNullException("response");
			}
			Match match = Regex.Match(response, "<.+>");
			if (match.Success)
			{
				ApopTimeStamp = match.Value;
				ApopSupported = true;
			}
		}

		private static void IsOkResponse(string response)
		{
			if (response == null)
			{
				throw new PopServerException("The stream used to retrieve responses from was closed");
			}
			if (response.StartsWith("+", StringComparison.OrdinalIgnoreCase))
			{
				return;
			}
			throw new PopServerException("The server did not respond with a + response. The response was: \"" + response + "\"");
		}

		private void SendCommand(string command)
		{
			byte[] bytes = Encoding.ASCII.GetBytes(command + "\r\n");
			DefaultLogger.Log.LogDebug($"SendCommand: \"{command}\"");
			Stream.Write(bytes, 0, bytes.Length);
			Stream.Flush();
			LastServerResponse = StreamUtility.ReadLineAsAscii(Stream);
			DefaultLogger.Log.LogDebug($"Server-Response: \"{LastServerResponse}\"");
			IsOkResponse(LastServerResponse);
		}

		private int SendCommandIntResponse(string command, int location)
		{
			SendCommand(command);
			return int.Parse(LastServerResponse.Split(' ')[location], CultureInfo.InvariantCulture);
		}

		private byte[] GetMessageAsBytes(int messageNumber, bool askOnlyForHeaders)
		{
			AssertDisposed();
			ValidateMessageNumber(messageNumber);
			if (State != ConnectionState.Transaction)
			{
				throw new InvalidUseException("Cannot fetch a message, when the user has not been authenticated yet");
			}
			if (askOnlyForHeaders)
			{
				SendCommand("TOP " + messageNumber + " 0");
			}
			else
			{
				SendCommand("RETR " + messageNumber);
			}
			using (MemoryStream memoryStream = new MemoryStream())
			{
				bool flag = true;
				byte[] array;
				while (!IsLastLineInMultiLineResponse(array = StreamUtility.ReadLineAsBytes(Stream)))
				{
					if (!flag)
					{
						byte[] bytes = Encoding.ASCII.GetBytes("\r\n");
						memoryStream.Write(bytes, 0, bytes.Length);
					}
					else
					{
						flag = false;
					}
					if (array.Length > 0 && array[0] == 46)
					{
						memoryStream.Write(array, 1, array.Length - 1);
					}
					else
					{
						memoryStream.Write(array, 0, array.Length);
					}
				}
				if (askOnlyForHeaders)
				{
					byte[] bytes2 = Encoding.ASCII.GetBytes("\r\n");
					memoryStream.Write(bytes2, 0, bytes2.Length);
				}
				return memoryStream.ToArray();
			}
		}

		private static bool IsLastLineInMultiLineResponse(byte[] bytesReceived)
		{
			if (bytesReceived == null)
			{
				throw new ArgumentNullException("bytesReceived");
			}
			if (bytesReceived.Length == 1)
			{
				return bytesReceived[0] == 46;
			}
			return false;
		}

		private static bool IsLastLineInMultiLineResponse(string lineReceived)
		{
			if (lineReceived == null)
			{
				throw new ArgumentNullException("lineReceived");
			}
			if (lineReceived.Length == 1)
			{
				return IsLastLineInMultiLineResponse(Encoding.ASCII.GetBytes(lineReceived));
			}
			return false;
		}

		private static void ValidateMessageNumber(int messageNumber)
		{
			if (messageNumber <= 0)
			{
				throw new InvalidUseException("The messageNumber argument cannot have a value of zero or less. Valid messageNumber is in the range [1, messageCount]");
			}
		}

		private void DisconnectStreams()
		{
			try
			{
				Stream.Close();
			}
			finally
			{
				SetInitialValues();
			}
		}

		private void SetInitialValues()
		{
			ApopTimeStamp = null;
			Connected = false;
			State = ConnectionState.Disconnected;
			ApopSupported = false;
		}

		private static void CheckFailedLoginServerResponse(string serverErrorResponse, PopServerException e)
		{
			string text = serverErrorResponse.ToUpperInvariant();
			if (text.Contains("[IN-USE]") || text.Contains("LOCK"))
			{
				DefaultLogger.Log.LogError("Authentication: maildrop is locked or in-use");
				throw new PopServerLockedException(e);
			}
			if (text.Contains("[LOGIN-DELAY]"))
			{
				throw new LoginDelayException(e);
			}
		}
	}
}
