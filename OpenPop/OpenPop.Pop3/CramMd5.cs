using System;
using System.Security.Cryptography;
using System.Text;

namespace OpenPop.Pop3
{
	internal static class CramMd5
	{
		private static readonly byte[] ipad;

		private static readonly byte[] opad;

		static CramMd5()
		{
			ipad = new byte[64];
			opad = new byte[64];
			for (int i = 0; i < ipad.Length; i++)
			{
				ipad[i] = 54;
				opad[i] = 92;
			}
		}

		internal static string ComputeDigest(string username, string password, string challenge)
		{
			if (username == null)
			{
				throw new ArgumentNullException("username");
			}
			if (password == null)
			{
				throw new ArgumentNullException("password");
			}
			if (challenge == null)
			{
				throw new ArgumentNullException("challenge");
			}
			byte[] sharedSecretInBytes = GetSharedSecretInBytes(password);
			byte[] two = Convert.FromBase64String(challenge);
			byte[] one = Xor(sharedSecretInBytes, opad);
			byte[] one2 = Xor(sharedSecretInBytes, ipad);
			byte[] value = Hash(Concatenate(one, Hash(Concatenate(one2, two))));
			string str = BitConverter.ToString(value).Replace("-", "").ToLowerInvariant();
			return Convert.ToBase64String(Encoding.ASCII.GetBytes(username + " " + str));
		}

		private static byte[] Hash(byte[] toHash)
		{
			if (toHash == null)
			{
				throw new ArgumentNullException("toHash");
			}
			using (MD5 mD = new MD5CryptoServiceProvider())
			{
				return mD.ComputeHash(toHash);
			}
		}

		private static byte[] Concatenate(byte[] one, byte[] two)
		{
			if (one == null)
			{
				throw new ArgumentNullException("one");
			}
			if (two == null)
			{
				throw new ArgumentNullException("two");
			}
			byte[] array = new byte[one.Length + two.Length];
			Buffer.BlockCopy(one, 0, array, 0, one.Length);
			Buffer.BlockCopy(two, 0, array, one.Length, two.Length);
			return array;
		}

		private static byte[] Xor(byte[] toXor, byte[] toXorWith)
		{
			if (toXor == null)
			{
				throw new ArgumentNullException("toXor");
			}
			if (toXorWith == null)
			{
				throw new ArgumentNullException("toXorWith");
			}
			if (toXor.Length != toXorWith.Length)
			{
				throw new ArgumentException("The lengths of the arrays must be equal");
			}
			byte[] array = new byte[toXor.Length];
			for (int i = 0; i < toXor.Length; i++)
			{
				array[i] = toXor[i];
				array[i] ^= toXorWith[i];
			}
			return array;
		}

		private static byte[] GetSharedSecretInBytes(string password)
		{
			if (password == null)
			{
				throw new ArgumentNullException("password");
			}
			byte[] array = Encoding.ASCII.GetBytes(password);
			if (array.Length > 64)
			{
				array = new MD5CryptoServiceProvider().ComputeHash(array);
			}
			if (array.Length != 64)
			{
				byte[] array2 = new byte[64];
				for (int i = 0; i < array.Length; i++)
				{
					array2[i] = array[i];
				}
				return array2;
			}
			return array;
		}
	}
}
