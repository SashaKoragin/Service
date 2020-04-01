using System;
using System.Security.Cryptography;
using System.Text;

namespace OpenPop.Pop3
{
	internal static class Apop
	{
		public static string ComputeDigest(string password, string serverTimestamp)
		{
			if (password == null)
			{
				throw new ArgumentNullException("password");
			}
			if (serverTimestamp == null)
			{
				throw new ArgumentNullException("serverTimestamp");
			}
			byte[] bytes = Encoding.ASCII.GetBytes(serverTimestamp + password);
			using (MD5 mD = new MD5CryptoServiceProvider())
			{
				byte[] value = mD.ComputeHash(bytes);
				return BitConverter.ToString(value).Replace("-", "").ToLowerInvariant();
			}
		}
	}
}
